
# <a name="userprofile"></a><span data-ttu-id="c4dab-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="c4dab-101">userProfile</span></span>

### <span data-ttu-id="c4dab-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="c4dab-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4dab-104">Требования</span><span class="sxs-lookup"><span data-stu-id="c4dab-104">Requirements</span></span>

|<span data-ttu-id="c4dab-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="c4dab-105">Requirement</span></span>| <span data-ttu-id="c4dab-106">Значение</span><span class="sxs-lookup"><span data-stu-id="c4dab-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4dab-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c4dab-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4dab-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c4dab-108">1.0</span></span>|
|[<span data-ttu-id="c4dab-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c4dab-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4dab-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4dab-110">ReadItem</span></span>|
|[<span data-ttu-id="c4dab-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c4dab-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4dab-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c4dab-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c4dab-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="c4dab-113">Members and methods</span></span>

| <span data-ttu-id="c4dab-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="c4dab-114">Member</span></span> | <span data-ttu-id="c4dab-115">Тип</span><span class="sxs-lookup"><span data-stu-id="c4dab-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c4dab-116">accountType</span><span class="sxs-lookup"><span data-stu-id="c4dab-116">AccountType</span></span>](#accounttype-string) | <span data-ttu-id="c4dab-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="c4dab-117">Member</span></span> |
| [<span data-ttu-id="c4dab-118">displayName</span><span class="sxs-lookup"><span data-stu-id="c4dab-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="c4dab-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="c4dab-119">Member</span></span> |
| [<span data-ttu-id="c4dab-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="c4dab-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="c4dab-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="c4dab-121">Member</span></span> |
| [<span data-ttu-id="c4dab-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="c4dab-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="c4dab-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="c4dab-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c4dab-124">Элементы</span><span class="sxs-lookup"><span data-stu-id="c4dab-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="c4dab-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="c4dab-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="c4dab-126">В настоящее время этот элемент поддерживается только в Outlook 2016 для Mac (сборка 16.9.1212 или более поздняя версия).</span><span class="sxs-lookup"><span data-stu-id="c4dab-126">This member is currently only supported in Outlook 2016 for Mac, build 16.9.1212 and greater.</span></span>

<span data-ttu-id="c4dab-127">Возвращает тип учетной записи пользователя, связанной с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="c4dab-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="c4dab-128">Возможные значения перечислены в таблице ниже.</span><span class="sxs-lookup"><span data-stu-id="c4dab-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="c4dab-129">Значение</span><span class="sxs-lookup"><span data-stu-id="c4dab-129">Value</span></span> | <span data-ttu-id="c4dab-130">Описание</span><span class="sxs-lookup"><span data-stu-id="c4dab-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="c4dab-131">Почтовый ящик размещен на локальном сервере Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="c4dab-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="c4dab-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="c4dab-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="c4dab-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="c4dab-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="c4dab-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="c4dab-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="c4dab-135">Тип:</span><span class="sxs-lookup"><span data-stu-id="c4dab-135">Type:</span></span>

*   <span data-ttu-id="c4dab-136">String</span><span class="sxs-lookup"><span data-stu-id="c4dab-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4dab-137">Требования</span><span class="sxs-lookup"><span data-stu-id="c4dab-137">Requirements</span></span>

|<span data-ttu-id="c4dab-138">Requirement</span><span class="sxs-lookup"><span data-stu-id="c4dab-138">Requirement</span></span>| <span data-ttu-id="c4dab-139">Значение</span><span class="sxs-lookup"><span data-stu-id="c4dab-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4dab-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="c4dab-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4dab-141">1.6</span><span class="sxs-lookup"><span data-stu-id="c4dab-141">1.6</span></span> |
|[<span data-ttu-id="c4dab-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c4dab-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4dab-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4dab-143">ReadItem</span></span>|
|[<span data-ttu-id="c4dab-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c4dab-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4dab-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c4dab-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c4dab-146">Пример</span><span class="sxs-lookup"><span data-stu-id="c4dab-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="c4dab-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c4dab-147">displayName :String</span></span>

<span data-ttu-id="c4dab-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="c4dab-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c4dab-149">Тип:</span><span class="sxs-lookup"><span data-stu-id="c4dab-149">Type:</span></span>

*   <span data-ttu-id="c4dab-150">String</span><span class="sxs-lookup"><span data-stu-id="c4dab-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4dab-151">Требования</span><span class="sxs-lookup"><span data-stu-id="c4dab-151">Requirements</span></span>

|<span data-ttu-id="c4dab-152">Requirement</span><span class="sxs-lookup"><span data-stu-id="c4dab-152">Requirement</span></span>| <span data-ttu-id="c4dab-153">Значение</span><span class="sxs-lookup"><span data-stu-id="c4dab-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4dab-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c4dab-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4dab-155">1.0</span><span class="sxs-lookup"><span data-stu-id="c4dab-155">1.0</span></span>|
|[<span data-ttu-id="c4dab-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c4dab-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4dab-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4dab-157">ReadItem</span></span>|
|[<span data-ttu-id="c4dab-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c4dab-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4dab-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c4dab-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c4dab-160">Пример</span><span class="sxs-lookup"><span data-stu-id="c4dab-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c4dab-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c4dab-161">emailAddress :String</span></span>

<span data-ttu-id="c4dab-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="c4dab-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c4dab-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="c4dab-163">Type:</span></span>

*   <span data-ttu-id="c4dab-164">String</span><span class="sxs-lookup"><span data-stu-id="c4dab-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4dab-165">Требования</span><span class="sxs-lookup"><span data-stu-id="c4dab-165">Requirements</span></span>

|<span data-ttu-id="c4dab-166">Requirement</span><span class="sxs-lookup"><span data-stu-id="c4dab-166">Requirement</span></span>| <span data-ttu-id="c4dab-167">Значение</span><span class="sxs-lookup"><span data-stu-id="c4dab-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4dab-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c4dab-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4dab-169">1.0</span><span class="sxs-lookup"><span data-stu-id="c4dab-169">1.0</span></span>|
|[<span data-ttu-id="c4dab-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c4dab-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4dab-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4dab-171">ReadItem</span></span>|
|[<span data-ttu-id="c4dab-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c4dab-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4dab-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c4dab-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c4dab-174">Пример</span><span class="sxs-lookup"><span data-stu-id="c4dab-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c4dab-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c4dab-175">timeZone :String</span></span>

<span data-ttu-id="c4dab-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="c4dab-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c4dab-177">Тип:</span><span class="sxs-lookup"><span data-stu-id="c4dab-177">Type:</span></span>

*   <span data-ttu-id="c4dab-178">String</span><span class="sxs-lookup"><span data-stu-id="c4dab-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c4dab-179">Требования</span><span class="sxs-lookup"><span data-stu-id="c4dab-179">Requirements</span></span>

|<span data-ttu-id="c4dab-180">Requirement</span><span class="sxs-lookup"><span data-stu-id="c4dab-180">Requirement</span></span>| <span data-ttu-id="c4dab-181">Значение</span><span class="sxs-lookup"><span data-stu-id="c4dab-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="c4dab-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c4dab-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c4dab-183">1.0</span><span class="sxs-lookup"><span data-stu-id="c4dab-183">1.0</span></span>|
|[<span data-ttu-id="c4dab-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c4dab-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c4dab-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c4dab-185">ReadItem</span></span>|
|[<span data-ttu-id="c4dab-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c4dab-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c4dab-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c4dab-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c4dab-188">Пример</span><span class="sxs-lookup"><span data-stu-id="c4dab-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```