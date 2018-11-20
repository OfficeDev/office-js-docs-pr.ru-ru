
# <a name="userprofile"></a><span data-ttu-id="ee5bf-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="ee5bf-101">userProfile</span></span>

### <span data-ttu-id="ee5bf-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="ee5bf-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5bf-104">Требования</span><span class="sxs-lookup"><span data-stu-id="ee5bf-104">Requirements</span></span>

|<span data-ttu-id="ee5bf-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="ee5bf-105">Requirement</span></span>| <span data-ttu-id="ee5bf-106">Значение</span><span class="sxs-lookup"><span data-stu-id="ee5bf-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5bf-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ee5bf-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee5bf-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ee5bf-108">1.0</span></span>|
|[<span data-ttu-id="ee5bf-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ee5bf-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee5bf-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee5bf-110">ReadItem</span></span>|
|[<span data-ttu-id="ee5bf-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ee5bf-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee5bf-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ee5bf-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ee5bf-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="ee5bf-113">Members and methods</span></span>

| <span data-ttu-id="ee5bf-114">Член</span><span class="sxs-lookup"><span data-stu-id="ee5bf-114">Member</span></span> | <span data-ttu-id="ee5bf-115">Type</span><span class="sxs-lookup"><span data-stu-id="ee5bf-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ee5bf-116">accountType</span><span class="sxs-lookup"><span data-stu-id="ee5bf-116">AccountType</span></span>](#accounttype-string) | <span data-ttu-id="ee5bf-117">Member</span><span class="sxs-lookup"><span data-stu-id="ee5bf-117">Member</span></span> |
| [<span data-ttu-id="ee5bf-118">displayName</span><span class="sxs-lookup"><span data-stu-id="ee5bf-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="ee5bf-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="ee5bf-119">Member</span></span> |
| [<span data-ttu-id="ee5bf-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="ee5bf-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="ee5bf-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="ee5bf-121">Member</span></span> |
| [<span data-ttu-id="ee5bf-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="ee5bf-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="ee5bf-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="ee5bf-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="ee5bf-124">Members</span><span class="sxs-lookup"><span data-stu-id="ee5bf-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="ee5bf-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="ee5bf-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="ee5bf-126">В настоящее время этот элемент поддерживается только в Outlook 2016 или более поздней версии для Mac (сборка 16.9.1212 или более поздняя версия).</span><span class="sxs-lookup"><span data-stu-id="ee5bf-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="ee5bf-127">Возвращает тип учетной записи пользователя, связанной с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="ee5bf-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="ee5bf-128">Возможные значения перечислены в таблице ниже.</span><span class="sxs-lookup"><span data-stu-id="ee5bf-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="ee5bf-129">Значение</span><span class="sxs-lookup"><span data-stu-id="ee5bf-129">Value</span></span> | <span data-ttu-id="ee5bf-130">Описание</span><span class="sxs-lookup"><span data-stu-id="ee5bf-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="ee5bf-131">Почтовый ящик размещен на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="ee5bf-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="ee5bf-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="ee5bf-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="ee5bf-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="ee5bf-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="ee5bf-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="ee5bf-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="ee5bf-135">Тип:</span><span class="sxs-lookup"><span data-stu-id="ee5bf-135">Type:</span></span>

*   <span data-ttu-id="ee5bf-136">String</span><span class="sxs-lookup"><span data-stu-id="ee5bf-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5bf-137">Требования</span><span class="sxs-lookup"><span data-stu-id="ee5bf-137">Requirements</span></span>

|<span data-ttu-id="ee5bf-138">Requirement</span><span class="sxs-lookup"><span data-stu-id="ee5bf-138">Requirement</span></span>| <span data-ttu-id="ee5bf-139">Значение</span><span class="sxs-lookup"><span data-stu-id="ee5bf-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5bf-140">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ee5bf-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee5bf-141">1.6</span><span class="sxs-lookup"><span data-stu-id="ee5bf-141">1.6</span></span> |
|[<span data-ttu-id="ee5bf-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ee5bf-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee5bf-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee5bf-143">ReadItem</span></span>|
|[<span data-ttu-id="ee5bf-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ee5bf-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee5bf-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ee5bf-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5bf-146">Пример</span><span class="sxs-lookup"><span data-stu-id="ee5bf-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="ee5bf-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ee5bf-147">displayName :String</span></span>

<span data-ttu-id="ee5bf-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="ee5bf-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5bf-149">Тип:</span><span class="sxs-lookup"><span data-stu-id="ee5bf-149">Type:</span></span>

*   <span data-ttu-id="ee5bf-150">String</span><span class="sxs-lookup"><span data-stu-id="ee5bf-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5bf-151">Требования</span><span class="sxs-lookup"><span data-stu-id="ee5bf-151">Requirements</span></span>

|<span data-ttu-id="ee5bf-152">Requirement</span><span class="sxs-lookup"><span data-stu-id="ee5bf-152">Requirement</span></span>| <span data-ttu-id="ee5bf-153">Значение</span><span class="sxs-lookup"><span data-stu-id="ee5bf-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5bf-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ee5bf-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee5bf-155">1.0</span><span class="sxs-lookup"><span data-stu-id="ee5bf-155">1.0</span></span>|
|[<span data-ttu-id="ee5bf-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ee5bf-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee5bf-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee5bf-157">ReadItem</span></span>|
|[<span data-ttu-id="ee5bf-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ee5bf-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee5bf-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ee5bf-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5bf-160">Пример</span><span class="sxs-lookup"><span data-stu-id="ee5bf-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ee5bf-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ee5bf-161">emailAddress :String</span></span>

<span data-ttu-id="ee5bf-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="ee5bf-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5bf-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="ee5bf-163">Type:</span></span>

*   <span data-ttu-id="ee5bf-164">String</span><span class="sxs-lookup"><span data-stu-id="ee5bf-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5bf-165">Требования</span><span class="sxs-lookup"><span data-stu-id="ee5bf-165">Requirements</span></span>

|<span data-ttu-id="ee5bf-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="ee5bf-166">Requirement</span></span>| <span data-ttu-id="ee5bf-167">Значение</span><span class="sxs-lookup"><span data-stu-id="ee5bf-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5bf-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ee5bf-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee5bf-169">1.0</span><span class="sxs-lookup"><span data-stu-id="ee5bf-169">1.0</span></span>|
|[<span data-ttu-id="ee5bf-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ee5bf-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee5bf-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee5bf-171">ReadItem</span></span>|
|[<span data-ttu-id="ee5bf-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ee5bf-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee5bf-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ee5bf-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5bf-174">Пример</span><span class="sxs-lookup"><span data-stu-id="ee5bf-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ee5bf-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ee5bf-175">timeZone :String</span></span>

<span data-ttu-id="ee5bf-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ee5bf-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ee5bf-177">Тип:</span><span class="sxs-lookup"><span data-stu-id="ee5bf-177">Type:</span></span>

*   <span data-ttu-id="ee5bf-178">String</span><span class="sxs-lookup"><span data-stu-id="ee5bf-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee5bf-179">Требования</span><span class="sxs-lookup"><span data-stu-id="ee5bf-179">Requirements</span></span>

|<span data-ttu-id="ee5bf-180">Requirement</span><span class="sxs-lookup"><span data-stu-id="ee5bf-180">Requirement</span></span>| <span data-ttu-id="ee5bf-181">Значение</span><span class="sxs-lookup"><span data-stu-id="ee5bf-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee5bf-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ee5bf-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee5bf-183">1.0</span><span class="sxs-lookup"><span data-stu-id="ee5bf-183">1.0</span></span>|
|[<span data-ttu-id="ee5bf-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ee5bf-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee5bf-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee5bf-185">ReadItem</span></span>|
|[<span data-ttu-id="ee5bf-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ee5bf-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee5bf-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ee5bf-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee5bf-188">Пример</span><span class="sxs-lookup"><span data-stu-id="ee5bf-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```