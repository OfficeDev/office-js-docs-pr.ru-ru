
# <a name="diagnostics"></a><span data-ttu-id="14cba-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="14cba-101">diagnostics</span></span>

### <span data-ttu-id="14cba-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="14cba-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="14cba-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="14cba-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14cba-105">Требования</span><span class="sxs-lookup"><span data-stu-id="14cba-105">Requirements</span></span>

|<span data-ttu-id="14cba-106">Требование</span><span class="sxs-lookup"><span data-stu-id="14cba-106">Requirement</span></span>| <span data-ttu-id="14cba-107">Значение</span><span class="sxs-lookup"><span data-stu-id="14cba-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="14cba-108">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="14cba-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14cba-109">1.0</span><span class="sxs-lookup"><span data-stu-id="14cba-109">1.0</span></span>|
|[<span data-ttu-id="14cba-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14cba-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14cba-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14cba-111">ReadItem</span></span>|
|[<span data-ttu-id="14cba-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14cba-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14cba-113">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="14cba-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="14cba-114">Члены</span><span class="sxs-lookup"><span data-stu-id="14cba-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="14cba-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="14cba-115">hostName :String</span></span>

<span data-ttu-id="14cba-116">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="14cba-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="14cba-117">Строка может иметь одно из следующих значений: `Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="14cba-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, , or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="14cba-118">Тип:</span><span class="sxs-lookup"><span data-stu-id="14cba-118">Type:</span></span>

*   <span data-ttu-id="14cba-119">Строка</span><span class="sxs-lookup"><span data-stu-id="14cba-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14cba-120">Требования</span><span class="sxs-lookup"><span data-stu-id="14cba-120">Requirements</span></span>

|<span data-ttu-id="14cba-121">Требование</span><span class="sxs-lookup"><span data-stu-id="14cba-121">Requirement</span></span>| <span data-ttu-id="14cba-122">Значение</span><span class="sxs-lookup"><span data-stu-id="14cba-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="14cba-123">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="14cba-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14cba-124">1.0</span><span class="sxs-lookup"><span data-stu-id="14cba-124">1.0</span></span>|
|[<span data-ttu-id="14cba-125">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14cba-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14cba-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14cba-126">ReadItem</span></span>|
|[<span data-ttu-id="14cba-127">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14cba-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14cba-128">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="14cba-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="14cba-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="14cba-129">hostVersion :String</span></span>

<span data-ttu-id="14cba-130">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="14cba-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="14cba-p102">Если почтовая надстройка запущена в классическом клиенте Outlook или Outlook для iOS, свойство `hostVersion` возвращает версию ведущего приложения — Outlook. В Outlook Web App это свойство возвращает версию Exchange Server. Например, строка `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="14cba-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="14cba-134">Тип:</span><span class="sxs-lookup"><span data-stu-id="14cba-134">Type:</span></span>

*   <span data-ttu-id="14cba-135">Строка</span><span class="sxs-lookup"><span data-stu-id="14cba-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14cba-136">Требования</span><span class="sxs-lookup"><span data-stu-id="14cba-136">Requirements</span></span>

|<span data-ttu-id="14cba-137">Требование</span><span class="sxs-lookup"><span data-stu-id="14cba-137">Requirement</span></span>| <span data-ttu-id="14cba-138">Значение</span><span class="sxs-lookup"><span data-stu-id="14cba-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="14cba-139">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="14cba-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14cba-140">1.0</span><span class="sxs-lookup"><span data-stu-id="14cba-140">1.0</span></span>|
|[<span data-ttu-id="14cba-141">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14cba-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14cba-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14cba-142">ReadItem</span></span>|
|[<span data-ttu-id="14cba-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14cba-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14cba-144">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="14cba-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="14cba-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="14cba-145">OWAView :String</span></span>

<span data-ttu-id="14cba-146">Получает строку, отображающую текущее представление Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="14cba-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="14cba-147">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="14cba-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="14cba-148">Если ведущее приложение — не Outlook Web App, тогда при получении доступа к этому свойству будет выдаваться значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="14cba-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="14cba-149">Outlook Web App включает три представления, которые соответствуют ширине экрана и окна, а также числу отображаемых столбцов:</span><span class="sxs-lookup"><span data-stu-id="14cba-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="14cba-p103">`OneColumn`, которое используется в случае узкого экрана. Outlook Web App использует этот макет размером в один столбец на экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="14cba-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="14cba-p104">`TwoColumns`, которое используется в случае более широкого экрана. Outlook Web App использует это представление на большинстве планшетных ПК.</span><span class="sxs-lookup"><span data-stu-id="14cba-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="14cba-p105">`ThreeColumns`, которое используется в случае широкого экрана. Например, Outlook Web App использует это представление в полноэкранном режиме на настольных компьютерах.</span><span class="sxs-lookup"><span data-stu-id="14cba-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="14cba-156">Тип:</span><span class="sxs-lookup"><span data-stu-id="14cba-156">Type:</span></span>

*   <span data-ttu-id="14cba-157">String</span><span class="sxs-lookup"><span data-stu-id="14cba-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14cba-158">Требования</span><span class="sxs-lookup"><span data-stu-id="14cba-158">Requirements</span></span>

|<span data-ttu-id="14cba-159">Требование</span><span class="sxs-lookup"><span data-stu-id="14cba-159">Requirement</span></span>| <span data-ttu-id="14cba-160">Значение</span><span class="sxs-lookup"><span data-stu-id="14cba-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="14cba-161">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="14cba-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14cba-162">1.0</span><span class="sxs-lookup"><span data-stu-id="14cba-162">1.0</span></span>|
|[<span data-ttu-id="14cba-163">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14cba-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14cba-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14cba-164">ReadItem</span></span>|
|[<span data-ttu-id="14cba-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14cba-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="14cba-166">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="14cba-166">Compose or read</span></span>|