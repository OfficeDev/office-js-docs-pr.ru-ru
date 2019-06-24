---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,7
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 2a79dbe7d392b809cf0de0b5ee7096473ea3e197
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127193"
---
# <a name="diagnostics"></a><span data-ttu-id="15897-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="15897-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="15897-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="15897-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="15897-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="15897-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="15897-105">Требования</span><span class="sxs-lookup"><span data-stu-id="15897-105">Requirements</span></span>

|<span data-ttu-id="15897-106">Требование</span><span class="sxs-lookup"><span data-stu-id="15897-106">Requirement</span></span>| <span data-ttu-id="15897-107">Значение</span><span class="sxs-lookup"><span data-stu-id="15897-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="15897-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="15897-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="15897-109">1.0</span><span class="sxs-lookup"><span data-stu-id="15897-109">1.0</span></span>|
|[<span data-ttu-id="15897-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="15897-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="15897-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="15897-111">ReadItem</span></span>|
|[<span data-ttu-id="15897-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="15897-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="15897-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="15897-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="15897-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="15897-114">Members and methods</span></span>

| <span data-ttu-id="15897-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="15897-115">Member</span></span> | <span data-ttu-id="15897-116">Тип</span><span class="sxs-lookup"><span data-stu-id="15897-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="15897-117">Сайту</span><span class="sxs-lookup"><span data-stu-id="15897-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="15897-118">Member</span><span class="sxs-lookup"><span data-stu-id="15897-118">Member</span></span> |
| [<span data-ttu-id="15897-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="15897-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="15897-120">Member</span><span class="sxs-lookup"><span data-stu-id="15897-120">Member</span></span> |
| [<span data-ttu-id="15897-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="15897-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="15897-122">Member</span><span class="sxs-lookup"><span data-stu-id="15897-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="15897-123">Members</span><span class="sxs-lookup"><span data-stu-id="15897-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="15897-124">Имя узла: строка</span><span class="sxs-lookup"><span data-stu-id="15897-124">hostName: String</span></span>

<span data-ttu-id="15897-125">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="15897-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="15897-126">Строка, которая может иметь одно из следующих значений: `Outlook`, `Mac Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="15897-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="15897-127">Тип</span><span class="sxs-lookup"><span data-stu-id="15897-127">Type</span></span>

*   <span data-ttu-id="15897-128">String</span><span class="sxs-lookup"><span data-stu-id="15897-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="15897-129">Требования</span><span class="sxs-lookup"><span data-stu-id="15897-129">Requirements</span></span>

|<span data-ttu-id="15897-130">Требование</span><span class="sxs-lookup"><span data-stu-id="15897-130">Requirement</span></span>| <span data-ttu-id="15897-131">Значение</span><span class="sxs-lookup"><span data-stu-id="15897-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="15897-132">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="15897-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="15897-133">1.0</span><span class="sxs-lookup"><span data-stu-id="15897-133">1.0</span></span>|
|[<span data-ttu-id="15897-134">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="15897-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="15897-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="15897-135">ReadItem</span></span>|
|[<span data-ttu-id="15897-136">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="15897-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="15897-137">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="15897-137">Compose or Read</span></span>|

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="15897-138">hostVersion: строка</span><span class="sxs-lookup"><span data-stu-id="15897-138">hostVersion: String</span></span>

<span data-ttu-id="15897-139">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="15897-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="15897-140">Если почтовая надстройка запущена на клиенте Outlook для настольных ПК или iOS `hostVersion` , свойство возвращает версию ведущего приложения, Outlook.</span><span class="sxs-lookup"><span data-stu-id="15897-140">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="15897-141">В Outlook в Интернете свойство возвращает версию сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="15897-141">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="15897-142">Пример — строка `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="15897-142">An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="15897-143">Тип</span><span class="sxs-lookup"><span data-stu-id="15897-143">Type</span></span>

*   <span data-ttu-id="15897-144">String</span><span class="sxs-lookup"><span data-stu-id="15897-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="15897-145">Требования</span><span class="sxs-lookup"><span data-stu-id="15897-145">Requirements</span></span>

|<span data-ttu-id="15897-146">Требование</span><span class="sxs-lookup"><span data-stu-id="15897-146">Requirement</span></span>| <span data-ttu-id="15897-147">Значение</span><span class="sxs-lookup"><span data-stu-id="15897-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="15897-148">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="15897-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="15897-149">1.0</span><span class="sxs-lookup"><span data-stu-id="15897-149">1.0</span></span>|
|[<span data-ttu-id="15897-150">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="15897-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="15897-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="15897-151">ReadItem</span></span>|
|[<span data-ttu-id="15897-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="15897-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="15897-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="15897-153">Compose or Read</span></span>|

---
---

#### <a name="owaview-string"></a><span data-ttu-id="15897-154">OWAView: строка</span><span class="sxs-lookup"><span data-stu-id="15897-154">OWAView: String</span></span>

<span data-ttu-id="15897-155">Получает строку, представляющую текущее представление Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="15897-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="15897-156">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="15897-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="15897-157">Если ведущее приложение не является Outlook в Интернете, то при доступе к этому свойству будет получен результат `undefined`.</span><span class="sxs-lookup"><span data-stu-id="15897-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="15897-158">В Outlook в Интернете есть три представления, которые соответствуют ширине экрана и окна, а также количество отображаемых столбцов:</span><span class="sxs-lookup"><span data-stu-id="15897-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="15897-159">`OneColumn`, который отображается, когда экран сужается.</span><span class="sxs-lookup"><span data-stu-id="15897-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="15897-160">В Outlook в Интернете этот макет с одним столбцом используется на всем экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="15897-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="15897-161">`TwoColumns`, который отображается, когда экран расширяется.</span><span class="sxs-lookup"><span data-stu-id="15897-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="15897-162">Outlook в Интернете использует это представление на большинстве планшетов.</span><span class="sxs-lookup"><span data-stu-id="15897-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="15897-163">`ThreeColumns` используется для полноразмерных экранов.</span><span class="sxs-lookup"><span data-stu-id="15897-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="15897-164">Например, в Outlook в Интернете это представление используется в полноэкранном окне на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="15897-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="15897-165">Тип</span><span class="sxs-lookup"><span data-stu-id="15897-165">Type</span></span>

*   <span data-ttu-id="15897-166">String</span><span class="sxs-lookup"><span data-stu-id="15897-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="15897-167">Требования</span><span class="sxs-lookup"><span data-stu-id="15897-167">Requirements</span></span>

|<span data-ttu-id="15897-168">Требование</span><span class="sxs-lookup"><span data-stu-id="15897-168">Requirement</span></span>| <span data-ttu-id="15897-169">Значение</span><span class="sxs-lookup"><span data-stu-id="15897-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="15897-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="15897-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="15897-171">1.0</span><span class="sxs-lookup"><span data-stu-id="15897-171">1.0</span></span>|
|[<span data-ttu-id="15897-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="15897-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="15897-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="15897-173">ReadItem</span></span>|
|[<span data-ttu-id="15897-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="15897-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="15897-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="15897-175">Compose or Read</span></span>|
