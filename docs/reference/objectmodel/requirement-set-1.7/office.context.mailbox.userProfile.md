
# <a name="userprofile"></a><span data-ttu-id="66135-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="66135-101">userProfile</span></span>

### <span data-ttu-id="66135-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="66135-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="66135-104">Требования</span><span class="sxs-lookup"><span data-stu-id="66135-104">Requirements</span></span>

|<span data-ttu-id="66135-105">Требование</span><span class="sxs-lookup"><span data-stu-id="66135-105">Requirement</span></span>| <span data-ttu-id="66135-106">Значение</span><span class="sxs-lookup"><span data-stu-id="66135-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="66135-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="66135-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66135-108">1.0</span><span class="sxs-lookup"><span data-stu-id="66135-108">1.0</span></span>|
|[<span data-ttu-id="66135-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="66135-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66135-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66135-110">ReadItem</span></span>|
|[<span data-ttu-id="66135-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="66135-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="66135-112">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="66135-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="66135-113">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="66135-113">Members and methods</span></span>

| <span data-ttu-id="66135-114">Член</span><span class="sxs-lookup"><span data-stu-id="66135-114">Member</span></span> | <span data-ttu-id="66135-115">Тип</span><span class="sxs-lookup"><span data-stu-id="66135-115">Type</span></span> |
|--------|------|
| <span data-ttu-id="66135-116">[AccountType](#accounttype-string)</span><span class="sxs-lookup"><span data-stu-id="66135-116">[](#accounttype-string)account_type=...</span></span> | <span data-ttu-id="66135-117">Член</span><span class="sxs-lookup"><span data-stu-id="66135-117">Member</span></span> |
| [<span data-ttu-id="66135-118">displayName</span><span class="sxs-lookup"><span data-stu-id="66135-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="66135-119">Член</span><span class="sxs-lookup"><span data-stu-id="66135-119">Member</span></span> |
| [<span data-ttu-id="66135-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="66135-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="66135-121">Член</span><span class="sxs-lookup"><span data-stu-id="66135-121">Member</span></span> |
| [<span data-ttu-id="66135-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="66135-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="66135-123">Член</span><span class="sxs-lookup"><span data-stu-id="66135-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="66135-124">Члены</span><span class="sxs-lookup"><span data-stu-id="66135-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="66135-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="66135-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="66135-126">Этот член в настоящее время поддерживается только в Outlook 2016 для Mac сборки 16.9.1212 и более поздних.</span><span class="sxs-lookup"><span data-stu-id="66135-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="66135-127">Получает тип учетной записи пользователя, связанной с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="66135-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="66135-128">В следующей таблице перечислены возможные значения.</span><span class="sxs-lookup"><span data-stu-id="66135-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="66135-129">Значение</span><span class="sxs-lookup"><span data-stu-id="66135-129">Value</span></span> | <span data-ttu-id="66135-130">Описание</span><span class="sxs-lookup"><span data-stu-id="66135-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="66135-131">Почтовый ящик находится на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="66135-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="66135-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="66135-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="66135-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="66135-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="66135-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="66135-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="66135-135">Тип:</span><span class="sxs-lookup"><span data-stu-id="66135-135">Type:</span></span>

*   <span data-ttu-id="66135-136">String</span><span class="sxs-lookup"><span data-stu-id="66135-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="66135-137">Требования</span><span class="sxs-lookup"><span data-stu-id="66135-137">Requirements</span></span>

|<span data-ttu-id="66135-138">Требование</span><span class="sxs-lookup"><span data-stu-id="66135-138">Requirement</span></span>| <span data-ttu-id="66135-139">Значение</span><span class="sxs-lookup"><span data-stu-id="66135-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="66135-140">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="66135-140">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66135-141">1.6</span><span class="sxs-lookup"><span data-stu-id="66135-141">1.6</span></span> |
|[<span data-ttu-id="66135-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="66135-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66135-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66135-143">ReadItem</span></span>|
|[<span data-ttu-id="66135-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="66135-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="66135-145">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="66135-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="66135-146">Пример</span><span class="sxs-lookup"><span data-stu-id="66135-146">Example</span></span>

```
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="66135-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="66135-147">displayName :String</span></span>

<span data-ttu-id="66135-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="66135-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="66135-149">Тип:</span><span class="sxs-lookup"><span data-stu-id="66135-149">Type:</span></span>

*   <span data-ttu-id="66135-150">String</span><span class="sxs-lookup"><span data-stu-id="66135-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="66135-151">Требования</span><span class="sxs-lookup"><span data-stu-id="66135-151">Requirements</span></span>

|<span data-ttu-id="66135-152">Требование</span><span class="sxs-lookup"><span data-stu-id="66135-152">Requirement</span></span>| <span data-ttu-id="66135-153">Значение</span><span class="sxs-lookup"><span data-stu-id="66135-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="66135-154">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="66135-154">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66135-155">1.0</span><span class="sxs-lookup"><span data-stu-id="66135-155">1.0</span></span>|
|[<span data-ttu-id="66135-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="66135-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66135-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66135-157">ReadItem</span></span>|
|[<span data-ttu-id="66135-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="66135-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="66135-159">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="66135-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="66135-160">Пример</span><span class="sxs-lookup"><span data-stu-id="66135-160">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="66135-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="66135-161">emailAddress :String</span></span>

<span data-ttu-id="66135-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="66135-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="66135-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="66135-163">Type:</span></span>

*   <span data-ttu-id="66135-164">String</span><span class="sxs-lookup"><span data-stu-id="66135-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="66135-165">Требования</span><span class="sxs-lookup"><span data-stu-id="66135-165">Requirements</span></span>

|<span data-ttu-id="66135-166">Требование</span><span class="sxs-lookup"><span data-stu-id="66135-166">Requirement</span></span>| <span data-ttu-id="66135-167">Значение</span><span class="sxs-lookup"><span data-stu-id="66135-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="66135-168">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="66135-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66135-169">1.0</span><span class="sxs-lookup"><span data-stu-id="66135-169">1.0</span></span>|
|[<span data-ttu-id="66135-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="66135-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66135-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66135-171">ReadItem</span></span>|
|[<span data-ttu-id="66135-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="66135-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="66135-173">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="66135-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="66135-174">Пример</span><span class="sxs-lookup"><span data-stu-id="66135-174">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="66135-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="66135-175">timeZone :String</span></span>

<span data-ttu-id="66135-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="66135-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="66135-177">Тип:</span><span class="sxs-lookup"><span data-stu-id="66135-177">Type:</span></span>

*   <span data-ttu-id="66135-178">String</span><span class="sxs-lookup"><span data-stu-id="66135-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="66135-179">Требования</span><span class="sxs-lookup"><span data-stu-id="66135-179">Requirements</span></span>

|<span data-ttu-id="66135-180">Требование</span><span class="sxs-lookup"><span data-stu-id="66135-180">Requirement</span></span>| <span data-ttu-id="66135-181">Значение</span><span class="sxs-lookup"><span data-stu-id="66135-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="66135-182">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="66135-182">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="66135-183">1.0</span><span class="sxs-lookup"><span data-stu-id="66135-183">1.0</span></span>|
|[<span data-ttu-id="66135-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="66135-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="66135-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="66135-185">ReadItem</span></span>|
|[<span data-ttu-id="66135-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="66135-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="66135-187">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="66135-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="66135-188">Пример</span><span class="sxs-lookup"><span data-stu-id="66135-188">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```