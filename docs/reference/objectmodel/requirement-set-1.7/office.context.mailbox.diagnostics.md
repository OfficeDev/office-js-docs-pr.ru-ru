---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 967834ff254f1b10d7518a012410beb2f327be68
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450361"
---
# <a name="diagnostics"></a><span data-ttu-id="9ee64-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="9ee64-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="9ee64-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="9ee64-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="9ee64-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="9ee64-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9ee64-105">Требования</span><span class="sxs-lookup"><span data-stu-id="9ee64-105">Requirements</span></span>

|<span data-ttu-id="9ee64-106">Требование</span><span class="sxs-lookup"><span data-stu-id="9ee64-106">Requirement</span></span>| <span data-ttu-id="9ee64-107">Значение</span><span class="sxs-lookup"><span data-stu-id="9ee64-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="9ee64-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9ee64-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9ee64-109">1.0</span><span class="sxs-lookup"><span data-stu-id="9ee64-109">1.0</span></span>|
|[<span data-ttu-id="9ee64-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9ee64-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9ee64-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9ee64-111">ReadItem</span></span>|
|[<span data-ttu-id="9ee64-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9ee64-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9ee64-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9ee64-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9ee64-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="9ee64-114">Members and methods</span></span>

| <span data-ttu-id="9ee64-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="9ee64-115">Member</span></span> | <span data-ttu-id="9ee64-116">Тип</span><span class="sxs-lookup"><span data-stu-id="9ee64-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9ee64-117">Сайту</span><span class="sxs-lookup"><span data-stu-id="9ee64-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="9ee64-118">Member</span><span class="sxs-lookup"><span data-stu-id="9ee64-118">Member</span></span> |
| [<span data-ttu-id="9ee64-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="9ee64-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="9ee64-120">Member</span><span class="sxs-lookup"><span data-stu-id="9ee64-120">Member</span></span> |
| [<span data-ttu-id="9ee64-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="9ee64-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="9ee64-122">Member</span><span class="sxs-lookup"><span data-stu-id="9ee64-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="9ee64-123">Элементы</span><span class="sxs-lookup"><span data-stu-id="9ee64-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="9ee64-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="9ee64-124">hostName :String</span></span>

<span data-ttu-id="9ee64-125">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="9ee64-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="9ee64-126">Строка, которая может иметь одно из следующих значений: `Outlook`, `Mac Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="9ee64-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="9ee64-127">Тип</span><span class="sxs-lookup"><span data-stu-id="9ee64-127">Type</span></span>

*   <span data-ttu-id="9ee64-128">String</span><span class="sxs-lookup"><span data-stu-id="9ee64-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9ee64-129">Требования</span><span class="sxs-lookup"><span data-stu-id="9ee64-129">Requirements</span></span>

|<span data-ttu-id="9ee64-130">Требование</span><span class="sxs-lookup"><span data-stu-id="9ee64-130">Requirement</span></span>| <span data-ttu-id="9ee64-131">Значение</span><span class="sxs-lookup"><span data-stu-id="9ee64-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="9ee64-132">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9ee64-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9ee64-133">1.0</span><span class="sxs-lookup"><span data-stu-id="9ee64-133">1.0</span></span>|
|[<span data-ttu-id="9ee64-134">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9ee64-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9ee64-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9ee64-135">ReadItem</span></span>|
|[<span data-ttu-id="9ee64-136">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9ee64-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9ee64-137">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9ee64-137">Compose or Read</span></span>|

---
---

####  <a name="hostversion-string"></a><span data-ttu-id="9ee64-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="9ee64-138">hostVersion :String</span></span>

<span data-ttu-id="9ee64-139">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="9ee64-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="9ee64-p101">Если почтовая надстройка запущена в классическом клиенте Outlook или Outlook для iOS, свойство `hostVersion` возвращает версию ведущего приложения, Outlook. В Outlook Web App это свойство возвращает версию Exchange Server. Пример — строка `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="9ee64-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="9ee64-143">Тип</span><span class="sxs-lookup"><span data-stu-id="9ee64-143">Type</span></span>

*   <span data-ttu-id="9ee64-144">String</span><span class="sxs-lookup"><span data-stu-id="9ee64-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9ee64-145">Требования</span><span class="sxs-lookup"><span data-stu-id="9ee64-145">Requirements</span></span>

|<span data-ttu-id="9ee64-146">Требование</span><span class="sxs-lookup"><span data-stu-id="9ee64-146">Requirement</span></span>| <span data-ttu-id="9ee64-147">Значение</span><span class="sxs-lookup"><span data-stu-id="9ee64-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="9ee64-148">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9ee64-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9ee64-149">1.0</span><span class="sxs-lookup"><span data-stu-id="9ee64-149">1.0</span></span>|
|[<span data-ttu-id="9ee64-150">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9ee64-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9ee64-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9ee64-151">ReadItem</span></span>|
|[<span data-ttu-id="9ee64-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9ee64-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9ee64-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9ee64-153">Compose or Read</span></span>|

---
---

####  <a name="owaview-string"></a><span data-ttu-id="9ee64-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="9ee64-154">OWAView :String</span></span>

<span data-ttu-id="9ee64-155">Получает строку, отображающую текущее представление Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="9ee64-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="9ee64-156">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="9ee64-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="9ee64-157">Если Outlook Web App — не ведущее приложение, при получении доступа к этому свойству будет выдаваться значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9ee64-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="9ee64-158">Outlook Web App включает три представления, которые соответствуют ширине экрана и окна, а также числу отображаемых столбцов.</span><span class="sxs-lookup"><span data-stu-id="9ee64-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="9ee64-p102">`OneColumn` используется в случае узкого экрана: Outlook Web App использует этот макет размером в один столбец на экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="9ee64-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="9ee64-p103">`TwoColumns` используется при более широком экране: Outlook Web App использует это представление на большинстве планшетных ПК.</span><span class="sxs-lookup"><span data-stu-id="9ee64-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="9ee64-p104">`ThreeColumns` используется для полноразмерных экранов. Например, Outlook Web App использует это представление в полноэкранном режиме на настольных компьютерах.</span><span class="sxs-lookup"><span data-stu-id="9ee64-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="9ee64-165">Тип</span><span class="sxs-lookup"><span data-stu-id="9ee64-165">Type</span></span>

*   <span data-ttu-id="9ee64-166">String</span><span class="sxs-lookup"><span data-stu-id="9ee64-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9ee64-167">Требования</span><span class="sxs-lookup"><span data-stu-id="9ee64-167">Requirements</span></span>

|<span data-ttu-id="9ee64-168">Требование</span><span class="sxs-lookup"><span data-stu-id="9ee64-168">Requirement</span></span>| <span data-ttu-id="9ee64-169">Значение</span><span class="sxs-lookup"><span data-stu-id="9ee64-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="9ee64-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9ee64-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9ee64-171">1.0</span><span class="sxs-lookup"><span data-stu-id="9ee64-171">1.0</span></span>|
|[<span data-ttu-id="9ee64-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9ee64-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9ee64-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9ee64-173">ReadItem</span></span>|
|[<span data-ttu-id="9ee64-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9ee64-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9ee64-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9ee64-175">Compose or Read</span></span>|
