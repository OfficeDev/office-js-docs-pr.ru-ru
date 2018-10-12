
# <a name="userprofile"></a><span data-ttu-id="051f4-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="051f4-101">userProfile</span></span>

### <span data-ttu-id="051f4-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="051f4-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="051f4-104">Требования</span><span class="sxs-lookup"><span data-stu-id="051f4-104">Requirements</span></span>

|<span data-ttu-id="051f4-105">Требование</span><span class="sxs-lookup"><span data-stu-id="051f4-105">Requirement</span></span>| <span data-ttu-id="051f4-106">Значение</span><span class="sxs-lookup"><span data-stu-id="051f4-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="051f4-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="051f4-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="051f4-108">1.0</span><span class="sxs-lookup"><span data-stu-id="051f4-108">1.0</span></span>|
|[<span data-ttu-id="051f4-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="051f4-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="051f4-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="051f4-110">ReadItem</span></span>|
|[<span data-ttu-id="051f4-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="051f4-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="051f4-112">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="051f4-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="051f4-113">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="051f4-113">Members and methods</span></span>

| <span data-ttu-id="051f4-114">Член</span><span class="sxs-lookup"><span data-stu-id="051f4-114">Member</span></span> | <span data-ttu-id="051f4-115">Тип</span><span class="sxs-lookup"><span data-stu-id="051f4-115">Type</span></span> |
|--------|------|
| <span data-ttu-id="051f4-116">[AccountType](#accounttype-string)</span><span class="sxs-lookup"><span data-stu-id="051f4-116">[](#accounttype-string)account_type=...</span></span> | <span data-ttu-id="051f4-117">Член</span><span class="sxs-lookup"><span data-stu-id="051f4-117">Member</span></span> |
| [<span data-ttu-id="051f4-118">displayName</span><span class="sxs-lookup"><span data-stu-id="051f4-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="051f4-119">Член</span><span class="sxs-lookup"><span data-stu-id="051f4-119">Member</span></span> |
| [<span data-ttu-id="051f4-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="051f4-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="051f4-121">Член</span><span class="sxs-lookup"><span data-stu-id="051f4-121">Member</span></span> |
| [<span data-ttu-id="051f4-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="051f4-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="051f4-123">Член</span><span class="sxs-lookup"><span data-stu-id="051f4-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="051f4-124">Члены</span><span class="sxs-lookup"><span data-stu-id="051f4-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="051f4-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="051f4-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="051f4-126">В настоящее время этот член поддерживается только в Outlook 2016 или более поздней версии для Mac (сборки 16.9.1212 или более поздней версии).</span><span class="sxs-lookup"><span data-stu-id="051f4-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="051f4-127">Получает тип учетной записи пользователя, связанной с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="051f4-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="051f4-128">В следующей таблице перечислены возможные значения.</span><span class="sxs-lookup"><span data-stu-id="051f4-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="051f4-129">Значение</span><span class="sxs-lookup"><span data-stu-id="051f4-129">Value</span></span> | <span data-ttu-id="051f4-130">Описание</span><span class="sxs-lookup"><span data-stu-id="051f4-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="051f4-131">Почтовый ящик находится на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="051f4-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="051f4-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="051f4-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="051f4-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="051f4-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="051f4-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="051f4-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="051f4-135">Тип:</span><span class="sxs-lookup"><span data-stu-id="051f4-135">Type:</span></span>

*   <span data-ttu-id="051f4-136">String</span><span class="sxs-lookup"><span data-stu-id="051f4-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="051f4-137">Требования</span><span class="sxs-lookup"><span data-stu-id="051f4-137">Requirements</span></span>

|<span data-ttu-id="051f4-138">Требование</span><span class="sxs-lookup"><span data-stu-id="051f4-138">Requirement</span></span>| <span data-ttu-id="051f4-139">Значение</span><span class="sxs-lookup"><span data-stu-id="051f4-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="051f4-140">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="051f4-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="051f4-141">1.6</span><span class="sxs-lookup"><span data-stu-id="051f4-141">1.6</span></span> |
|[<span data-ttu-id="051f4-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="051f4-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="051f4-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="051f4-143">ReadItem</span></span>|
|[<span data-ttu-id="051f4-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="051f4-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="051f4-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="051f4-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="051f4-146">Пример</span><span class="sxs-lookup"><span data-stu-id="051f4-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="051f4-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="051f4-147">displayName :String</span></span>

<span data-ttu-id="051f4-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="051f4-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="051f4-149">Тип:</span><span class="sxs-lookup"><span data-stu-id="051f4-149">Type:</span></span>

*   <span data-ttu-id="051f4-150">String</span><span class="sxs-lookup"><span data-stu-id="051f4-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="051f4-151">Требования</span><span class="sxs-lookup"><span data-stu-id="051f4-151">Requirements</span></span>

|<span data-ttu-id="051f4-152">Требование</span><span class="sxs-lookup"><span data-stu-id="051f4-152">Requirement</span></span>| <span data-ttu-id="051f4-153">Значение</span><span class="sxs-lookup"><span data-stu-id="051f4-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="051f4-154">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="051f4-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="051f4-155">1.0</span><span class="sxs-lookup"><span data-stu-id="051f4-155">1.0</span></span>|
|[<span data-ttu-id="051f4-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="051f4-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="051f4-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="051f4-157">ReadItem</span></span>|
|[<span data-ttu-id="051f4-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="051f4-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="051f4-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="051f4-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="051f4-160">Пример</span><span class="sxs-lookup"><span data-stu-id="051f4-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="051f4-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="051f4-161">emailAddress :String</span></span>

<span data-ttu-id="051f4-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="051f4-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="051f4-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="051f4-163">Type:</span></span>

*   <span data-ttu-id="051f4-164">String</span><span class="sxs-lookup"><span data-stu-id="051f4-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="051f4-165">Требования</span><span class="sxs-lookup"><span data-stu-id="051f4-165">Requirements</span></span>

|<span data-ttu-id="051f4-166">Требование</span><span class="sxs-lookup"><span data-stu-id="051f4-166">Requirement</span></span>| <span data-ttu-id="051f4-167">Значение</span><span class="sxs-lookup"><span data-stu-id="051f4-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="051f4-168">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="051f4-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="051f4-169">1.0</span><span class="sxs-lookup"><span data-stu-id="051f4-169">1.0</span></span>|
|[<span data-ttu-id="051f4-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="051f4-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="051f4-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="051f4-171">ReadItem</span></span>|
|[<span data-ttu-id="051f4-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="051f4-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="051f4-173">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="051f4-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="051f4-174">Пример</span><span class="sxs-lookup"><span data-stu-id="051f4-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="051f4-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="051f4-175">timeZone :String</span></span>

<span data-ttu-id="051f4-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="051f4-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="051f4-177">Тип:</span><span class="sxs-lookup"><span data-stu-id="051f4-177">Type:</span></span>

*   <span data-ttu-id="051f4-178">String</span><span class="sxs-lookup"><span data-stu-id="051f4-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="051f4-179">Требования</span><span class="sxs-lookup"><span data-stu-id="051f4-179">Requirements</span></span>

|<span data-ttu-id="051f4-180">Требование</span><span class="sxs-lookup"><span data-stu-id="051f4-180">Requirement</span></span>| <span data-ttu-id="051f4-181">Значение</span><span class="sxs-lookup"><span data-stu-id="051f4-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="051f4-182">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="051f4-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="051f4-183">1.0</span><span class="sxs-lookup"><span data-stu-id="051f4-183">1.0</span></span>|
|[<span data-ttu-id="051f4-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="051f4-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="051f4-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="051f4-185">ReadItem</span></span>|
|[<span data-ttu-id="051f4-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="051f4-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="051f4-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="051f4-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="051f4-188">Пример</span><span class="sxs-lookup"><span data-stu-id="051f4-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```