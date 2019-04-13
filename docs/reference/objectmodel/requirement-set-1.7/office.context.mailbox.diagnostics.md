---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 967834ff254f1b10d7518a012410beb2f327be68
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838468"
---
# <a name="diagnostics"></a><span data-ttu-id="67048-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="67048-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="67048-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="67048-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="67048-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="67048-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="67048-105">Требования</span><span class="sxs-lookup"><span data-stu-id="67048-105">Requirements</span></span>

|<span data-ttu-id="67048-106">Требование</span><span class="sxs-lookup"><span data-stu-id="67048-106">Requirement</span></span>| <span data-ttu-id="67048-107">Значение</span><span class="sxs-lookup"><span data-stu-id="67048-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="67048-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="67048-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67048-109">1.0</span><span class="sxs-lookup"><span data-stu-id="67048-109">1.0</span></span>|
|[<span data-ttu-id="67048-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="67048-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="67048-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67048-111">ReadItem</span></span>|
|[<span data-ttu-id="67048-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="67048-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="67048-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="67048-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="67048-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="67048-114">Members and methods</span></span>

| <span data-ttu-id="67048-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="67048-115">Member</span></span> | <span data-ttu-id="67048-116">Тип</span><span class="sxs-lookup"><span data-stu-id="67048-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="67048-117">Сайту</span><span class="sxs-lookup"><span data-stu-id="67048-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="67048-118">Member</span><span class="sxs-lookup"><span data-stu-id="67048-118">Member</span></span> |
| [<span data-ttu-id="67048-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="67048-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="67048-120">Member</span><span class="sxs-lookup"><span data-stu-id="67048-120">Member</span></span> |
| [<span data-ttu-id="67048-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="67048-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="67048-122">Member</span><span class="sxs-lookup"><span data-stu-id="67048-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="67048-123">Элементы</span><span class="sxs-lookup"><span data-stu-id="67048-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="67048-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="67048-124">hostName :String</span></span>

<span data-ttu-id="67048-125">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="67048-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="67048-126">Строка, которая может иметь одно из следующих значений: `Outlook`, `Mac Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="67048-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="67048-127">Тип</span><span class="sxs-lookup"><span data-stu-id="67048-127">Type</span></span>

*   <span data-ttu-id="67048-128">String</span><span class="sxs-lookup"><span data-stu-id="67048-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67048-129">Требования</span><span class="sxs-lookup"><span data-stu-id="67048-129">Requirements</span></span>

|<span data-ttu-id="67048-130">Требование</span><span class="sxs-lookup"><span data-stu-id="67048-130">Requirement</span></span>| <span data-ttu-id="67048-131">Значение</span><span class="sxs-lookup"><span data-stu-id="67048-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="67048-132">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="67048-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67048-133">1.0</span><span class="sxs-lookup"><span data-stu-id="67048-133">1.0</span></span>|
|[<span data-ttu-id="67048-134">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="67048-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="67048-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67048-135">ReadItem</span></span>|
|[<span data-ttu-id="67048-136">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="67048-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="67048-137">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="67048-137">Compose or Read</span></span>|

---
---

####  <a name="hostversion-string"></a><span data-ttu-id="67048-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="67048-138">hostVersion :String</span></span>

<span data-ttu-id="67048-139">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="67048-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="67048-p101">Если почтовая надстройка запущена в классическом клиенте Outlook или Outlook для iOS, свойство `hostVersion` возвращает версию ведущего приложения, Outlook. В Outlook Web App это свойство возвращает версию Exchange Server. Пример — строка `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="67048-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="67048-143">Тип</span><span class="sxs-lookup"><span data-stu-id="67048-143">Type</span></span>

*   <span data-ttu-id="67048-144">String</span><span class="sxs-lookup"><span data-stu-id="67048-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67048-145">Требования</span><span class="sxs-lookup"><span data-stu-id="67048-145">Requirements</span></span>

|<span data-ttu-id="67048-146">Требование</span><span class="sxs-lookup"><span data-stu-id="67048-146">Requirement</span></span>| <span data-ttu-id="67048-147">Значение</span><span class="sxs-lookup"><span data-stu-id="67048-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="67048-148">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="67048-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67048-149">1.0</span><span class="sxs-lookup"><span data-stu-id="67048-149">1.0</span></span>|
|[<span data-ttu-id="67048-150">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="67048-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="67048-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67048-151">ReadItem</span></span>|
|[<span data-ttu-id="67048-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="67048-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="67048-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="67048-153">Compose or Read</span></span>|

---
---

####  <a name="owaview-string"></a><span data-ttu-id="67048-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="67048-154">OWAView :String</span></span>

<span data-ttu-id="67048-155">Получает строку, отображающую текущее представление Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="67048-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="67048-156">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="67048-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="67048-157">Если Outlook Web App — не ведущее приложение, при получении доступа к этому свойству будет выдаваться значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="67048-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="67048-158">Outlook Web App включает три представления, которые соответствуют ширине экрана и окна, а также числу отображаемых столбцов.</span><span class="sxs-lookup"><span data-stu-id="67048-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="67048-p102">`OneColumn` используется в случае узкого экрана: Outlook Web App использует этот макет размером в один столбец на экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="67048-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="67048-p103">`TwoColumns` используется при более широком экране: Outlook Web App использует это представление на большинстве планшетных ПК.</span><span class="sxs-lookup"><span data-stu-id="67048-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="67048-p104">`ThreeColumns` используется для полноразмерных экранов. Например, Outlook Web App использует это представление в полноэкранном режиме на настольных компьютерах.</span><span class="sxs-lookup"><span data-stu-id="67048-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="67048-165">Тип</span><span class="sxs-lookup"><span data-stu-id="67048-165">Type</span></span>

*   <span data-ttu-id="67048-166">String</span><span class="sxs-lookup"><span data-stu-id="67048-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="67048-167">Требования</span><span class="sxs-lookup"><span data-stu-id="67048-167">Requirements</span></span>

|<span data-ttu-id="67048-168">Требование</span><span class="sxs-lookup"><span data-stu-id="67048-168">Requirement</span></span>| <span data-ttu-id="67048-169">Значение</span><span class="sxs-lookup"><span data-stu-id="67048-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="67048-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="67048-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="67048-171">1.0</span><span class="sxs-lookup"><span data-stu-id="67048-171">1.0</span></span>|
|[<span data-ttu-id="67048-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="67048-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="67048-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="67048-173">ReadItem</span></span>|
|[<span data-ttu-id="67048-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="67048-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="67048-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="67048-175">Compose or Read</span></span>|
