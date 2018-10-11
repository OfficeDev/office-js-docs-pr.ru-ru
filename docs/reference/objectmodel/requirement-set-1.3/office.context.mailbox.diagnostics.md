
# <a name="diagnostics"></a><span data-ttu-id="40aa4-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="40aa4-101">diagnostics</span></span>

### <span data-ttu-id="40aa4-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="40aa4-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="40aa4-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="40aa4-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="40aa4-105">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="40aa4-105">Requirements</span></span>

|<span data-ttu-id="40aa4-106">Требование</span><span class="sxs-lookup"><span data-stu-id="40aa4-106">Requirement</span></span>| <span data-ttu-id="40aa4-107">Значение</span><span class="sxs-lookup"><span data-stu-id="40aa4-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="40aa4-108">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="40aa4-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40aa4-109">1.0</span><span class="sxs-lookup"><span data-stu-id="40aa4-109">1.0</span></span>|
|[<span data-ttu-id="40aa4-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="40aa4-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40aa4-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40aa4-111">ReadItem</span></span>|
|[<span data-ttu-id="40aa4-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40aa4-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40aa4-113">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="40aa4-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="40aa4-114">Члены</span><span class="sxs-lookup"><span data-stu-id="40aa4-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="40aa4-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="40aa4-115">hostName :String</span></span>

<span data-ttu-id="40aa4-116">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="40aa4-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="40aa4-117">Строка может иметь одно из следующих значений: `Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="40aa4-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, , or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="40aa4-118">Тип:</span><span class="sxs-lookup"><span data-stu-id="40aa4-118">Type:</span></span>

*   <span data-ttu-id="40aa4-119">String</span><span class="sxs-lookup"><span data-stu-id="40aa4-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="40aa4-120">Требования</span><span class="sxs-lookup"><span data-stu-id="40aa4-120">Requirements</span></span>

|<span data-ttu-id="40aa4-121">Требование</span><span class="sxs-lookup"><span data-stu-id="40aa4-121">Requirement</span></span>| <span data-ttu-id="40aa4-122">Значение</span><span class="sxs-lookup"><span data-stu-id="40aa4-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="40aa4-123">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="40aa4-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40aa4-124">1.0</span><span class="sxs-lookup"><span data-stu-id="40aa4-124">1.0</span></span>|
|[<span data-ttu-id="40aa4-125">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="40aa4-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40aa4-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40aa4-126">ReadItem</span></span>|
|[<span data-ttu-id="40aa4-127">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40aa4-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40aa4-128">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="40aa4-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="40aa4-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="40aa4-129">hostVersion :String</span></span>

<span data-ttu-id="40aa4-130">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="40aa4-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="40aa4-p102">Если почтовая надстройка запущена в классическом клиенте Outlook или Outlook для iOS, свойство `hostVersion` возвращает версию ведущего приложения — Outlook. В Outlook Web App это свойство возвращает версию Exchange Server. Например, строка `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="40aa4-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="40aa4-134">Тип:</span><span class="sxs-lookup"><span data-stu-id="40aa4-134">Type:</span></span>

*   <span data-ttu-id="40aa4-135">String</span><span class="sxs-lookup"><span data-stu-id="40aa4-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="40aa4-136">Требования</span><span class="sxs-lookup"><span data-stu-id="40aa4-136">Requirements</span></span>

|<span data-ttu-id="40aa4-137">Требование</span><span class="sxs-lookup"><span data-stu-id="40aa4-137">Requirement</span></span>| <span data-ttu-id="40aa4-138">Значение</span><span class="sxs-lookup"><span data-stu-id="40aa4-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="40aa4-139">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="40aa4-139">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40aa4-140">1.0</span><span class="sxs-lookup"><span data-stu-id="40aa4-140">1.0</span></span>|
|[<span data-ttu-id="40aa4-141">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="40aa4-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40aa4-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40aa4-142">ReadItem</span></span>|
|[<span data-ttu-id="40aa4-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40aa4-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40aa4-144">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="40aa4-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="40aa4-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="40aa4-145">OWAView :String</span></span>

<span data-ttu-id="40aa4-146">Получает строку, отображающую текущее представление Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="40aa4-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="40aa4-147">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="40aa4-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="40aa4-148">Если ведущее приложение — не Outlook Web App, тогда при получении доступа к этому свойству будет выдаваться значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="40aa4-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="40aa4-149">Outlook Web App включает три представления, которые соответствуют ширине экрана и окна, а также числу отображаемых столбцов:</span><span class="sxs-lookup"><span data-stu-id="40aa4-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="40aa4-p103">`OneColumn`, которое используется в случае узкого экрана. Outlook Web App использует этот макет размером в один столбец на экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="40aa4-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="40aa4-p104">`TwoColumns`, которое используется в случае более широкого экрана. Outlook Web App использует это представление на большинстве планшетных ПК.</span><span class="sxs-lookup"><span data-stu-id="40aa4-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="40aa4-p105">`ThreeColumns`, которое используется в случае широкого экрана. Например, Outlook Web App использует это представление в полноэкранном режиме на настольных компьютерах.</span><span class="sxs-lookup"><span data-stu-id="40aa4-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="40aa4-156">Тип:</span><span class="sxs-lookup"><span data-stu-id="40aa4-156">Type:</span></span>

*   <span data-ttu-id="40aa4-157">String</span><span class="sxs-lookup"><span data-stu-id="40aa4-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="40aa4-158">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="40aa4-158">Requirements</span></span>

|<span data-ttu-id="40aa4-159">Требование</span><span class="sxs-lookup"><span data-stu-id="40aa4-159">Requirement</span></span>| <span data-ttu-id="40aa4-160">Значение</span><span class="sxs-lookup"><span data-stu-id="40aa4-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="40aa4-161">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="40aa4-161">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="40aa4-162">1.0</span><span class="sxs-lookup"><span data-stu-id="40aa4-162">1.0</span></span>|
|[<span data-ttu-id="40aa4-163">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="40aa4-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="40aa4-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="40aa4-164">ReadItem</span></span>|
|[<span data-ttu-id="40aa4-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40aa4-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="40aa4-166">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="40aa4-166">Compose or read</span></span>|