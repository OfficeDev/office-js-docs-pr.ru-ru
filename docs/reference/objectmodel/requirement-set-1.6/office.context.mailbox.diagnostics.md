
# <a name="diagnostics"></a><span data-ttu-id="cefbe-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="cefbe-101">diagnostics</span></span>

### <span data-ttu-id="cefbe-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="cefbe-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="cefbe-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="cefbe-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cefbe-105">Требования</span><span class="sxs-lookup"><span data-stu-id="cefbe-105">Requirements</span></span>

|<span data-ttu-id="cefbe-106">Требование</span><span class="sxs-lookup"><span data-stu-id="cefbe-106">Requirement</span></span>| <span data-ttu-id="cefbe-107">Значение</span><span class="sxs-lookup"><span data-stu-id="cefbe-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="cefbe-108">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="cefbe-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cefbe-109">1.0</span><span class="sxs-lookup"><span data-stu-id="cefbe-109">1.0</span></span>|
|[<span data-ttu-id="cefbe-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cefbe-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cefbe-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cefbe-111">ReadItem</span></span>|
|[<span data-ttu-id="cefbe-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cefbe-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cefbe-113">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="cefbe-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="cefbe-114">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="cefbe-114">Members and methods</span></span>

| <span data-ttu-id="cefbe-115">Член</span><span class="sxs-lookup"><span data-stu-id="cefbe-115">Member</span></span> | <span data-ttu-id="cefbe-116">Тип</span><span class="sxs-lookup"><span data-stu-id="cefbe-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="cefbe-117">hostName</span><span class="sxs-lookup"><span data-stu-id="cefbe-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="cefbe-118">Член</span><span class="sxs-lookup"><span data-stu-id="cefbe-118">Member</span></span> |
| [<span data-ttu-id="cefbe-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="cefbe-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="cefbe-120">Член</span><span class="sxs-lookup"><span data-stu-id="cefbe-120">Member</span></span> |
| [<span data-ttu-id="cefbe-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="cefbe-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="cefbe-122">Член</span><span class="sxs-lookup"><span data-stu-id="cefbe-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="cefbe-123">Члены</span><span class="sxs-lookup"><span data-stu-id="cefbe-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="cefbe-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="cefbe-124">hostName :String</span></span>

<span data-ttu-id="cefbe-125">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="cefbe-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="cefbe-126">Строка, которая может быть одним из следующих значений: `Outlook`, `Mac Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="cefbe-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="cefbe-127">Тип:</span><span class="sxs-lookup"><span data-stu-id="cefbe-127">Type:</span></span>

*   <span data-ttu-id="cefbe-128">String</span><span class="sxs-lookup"><span data-stu-id="cefbe-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cefbe-129">Требования</span><span class="sxs-lookup"><span data-stu-id="cefbe-129">Requirements</span></span>

|<span data-ttu-id="cefbe-130">Требование</span><span class="sxs-lookup"><span data-stu-id="cefbe-130">Requirement</span></span>| <span data-ttu-id="cefbe-131">Значение</span><span class="sxs-lookup"><span data-stu-id="cefbe-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="cefbe-132">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="cefbe-132">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cefbe-133">1.0</span><span class="sxs-lookup"><span data-stu-id="cefbe-133">1.0</span></span>|
|[<span data-ttu-id="cefbe-134">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cefbe-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cefbe-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cefbe-135">ReadItem</span></span>|
|[<span data-ttu-id="cefbe-136">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cefbe-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cefbe-137">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="cefbe-137">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="cefbe-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="cefbe-138">hostVersion :String</span></span>

<span data-ttu-id="cefbe-139">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="cefbe-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="cefbe-p102">Если почтовая надстройка запущена в классическом клиенте Outlook или Outlook для iOS, свойство `hostVersion` возвращает версию ведущего приложения — Outlook. В Outlook Web App это свойство возвращает версию Exchange Server. Например, строка `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="cefbe-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="cefbe-143">Тип:</span><span class="sxs-lookup"><span data-stu-id="cefbe-143">Type:</span></span>

*   <span data-ttu-id="cefbe-144">String</span><span class="sxs-lookup"><span data-stu-id="cefbe-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cefbe-145">Требования</span><span class="sxs-lookup"><span data-stu-id="cefbe-145">Requirements</span></span>

|<span data-ttu-id="cefbe-146">Требование</span><span class="sxs-lookup"><span data-stu-id="cefbe-146">Requirement</span></span>| <span data-ttu-id="cefbe-147">Значение</span><span class="sxs-lookup"><span data-stu-id="cefbe-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="cefbe-148">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="cefbe-148">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cefbe-149">1.0</span><span class="sxs-lookup"><span data-stu-id="cefbe-149">1.0</span></span>|
|[<span data-ttu-id="cefbe-150">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cefbe-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cefbe-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cefbe-151">ReadItem</span></span>|
|[<span data-ttu-id="cefbe-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cefbe-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cefbe-153">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="cefbe-153">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="cefbe-154">OWAView: String</span><span class="sxs-lookup"><span data-stu-id="cefbe-154">OWAView :String</span></span>

<span data-ttu-id="cefbe-155">Получает строку, отображающую текущее представление Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="cefbe-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="cefbe-156">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="cefbe-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="cefbe-157">Если Outlook Web App — не ведущее приложение, при получении доступа к этому свойству будет выдаваться значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="cefbe-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="cefbe-158">Outlook Web App включает три представления, которые соответствуют ширине экрана и окна, а также числу отображаемых столбцов.</span><span class="sxs-lookup"><span data-stu-id="cefbe-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="cefbe-p103">`OneColumn`используется в случае узкого экрана: Outlook Web App использует этот макет размером в один столбец на экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="cefbe-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="cefbe-p104">`TwoColumns`используется при более широком экране: Outlook Web App использует это представление на большинстве планшетных ПК.</span><span class="sxs-lookup"><span data-stu-id="cefbe-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="cefbe-p105">`ThreeColumns`используется для полноразмерных экранов. Например, Outlook Web App использует это представление в полноэкранном режиме на настольных компьютерах.</span><span class="sxs-lookup"><span data-stu-id="cefbe-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="cefbe-165">Тип:</span><span class="sxs-lookup"><span data-stu-id="cefbe-165">Type:</span></span>

*   <span data-ttu-id="cefbe-166">String</span><span class="sxs-lookup"><span data-stu-id="cefbe-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cefbe-167">Требования</span><span class="sxs-lookup"><span data-stu-id="cefbe-167">Requirements</span></span>

|<span data-ttu-id="cefbe-168">Требование</span><span class="sxs-lookup"><span data-stu-id="cefbe-168">Requirement</span></span>| <span data-ttu-id="cefbe-169">Значение</span><span class="sxs-lookup"><span data-stu-id="cefbe-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="cefbe-170">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="cefbe-170">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cefbe-171">1.0</span><span class="sxs-lookup"><span data-stu-id="cefbe-171">1.0</span></span>|
|[<span data-ttu-id="cefbe-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cefbe-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cefbe-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cefbe-173">ReadItem</span></span>|
|[<span data-ttu-id="cefbe-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cefbe-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="cefbe-175">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="cefbe-175">Compose or read</span></span>|