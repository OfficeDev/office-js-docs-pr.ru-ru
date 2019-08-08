---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,7
description: ''
ms.date: 08/05/2019
localization_priority: Normal
ms.openlocfilehash: e197374267d40056c6cb1dea8808e30f48eef65c
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231265"
---
# <a name="diagnostics"></a><span data-ttu-id="2ed73-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="2ed73-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="2ed73-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="2ed73-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="2ed73-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="2ed73-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ed73-105">Требования</span><span class="sxs-lookup"><span data-stu-id="2ed73-105">Requirements</span></span>

|<span data-ttu-id="2ed73-106">Требование</span><span class="sxs-lookup"><span data-stu-id="2ed73-106">Requirement</span></span>| <span data-ttu-id="2ed73-107">Значение</span><span class="sxs-lookup"><span data-stu-id="2ed73-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ed73-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2ed73-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ed73-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2ed73-109">1.0</span></span>|
|[<span data-ttu-id="2ed73-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2ed73-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ed73-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ed73-111">ReadItem</span></span>|
|[<span data-ttu-id="2ed73-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2ed73-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ed73-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2ed73-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2ed73-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="2ed73-114">Members and methods</span></span>

| <span data-ttu-id="2ed73-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="2ed73-115">Member</span></span> | <span data-ttu-id="2ed73-116">Тип</span><span class="sxs-lookup"><span data-stu-id="2ed73-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2ed73-117">Сайту</span><span class="sxs-lookup"><span data-stu-id="2ed73-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="2ed73-118">Member</span><span class="sxs-lookup"><span data-stu-id="2ed73-118">Member</span></span> |
| [<span data-ttu-id="2ed73-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="2ed73-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="2ed73-120">Member</span><span class="sxs-lookup"><span data-stu-id="2ed73-120">Member</span></span> |
| [<span data-ttu-id="2ed73-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="2ed73-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="2ed73-122">Member</span><span class="sxs-lookup"><span data-stu-id="2ed73-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="2ed73-123">Members</span><span class="sxs-lookup"><span data-stu-id="2ed73-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="2ed73-124">Имя узла: строка</span><span class="sxs-lookup"><span data-stu-id="2ed73-124">hostName: String</span></span>

<span data-ttu-id="2ed73-125">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="2ed73-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="2ed73-126">Строка, которая может иметь одно из следующих значений: `Outlook`, `OutlookWebApp`, `OutlookIOS` или `OutlookAndroid`.</span><span class="sxs-lookup"><span data-stu-id="2ed73-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

##### <a name="type"></a><span data-ttu-id="2ed73-127">Тип</span><span class="sxs-lookup"><span data-stu-id="2ed73-127">Type</span></span>

*   <span data-ttu-id="2ed73-128">String</span><span class="sxs-lookup"><span data-stu-id="2ed73-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ed73-129">Требования</span><span class="sxs-lookup"><span data-stu-id="2ed73-129">Requirements</span></span>

|<span data-ttu-id="2ed73-130">Требование</span><span class="sxs-lookup"><span data-stu-id="2ed73-130">Requirement</span></span>| <span data-ttu-id="2ed73-131">Значение</span><span class="sxs-lookup"><span data-stu-id="2ed73-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ed73-132">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2ed73-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ed73-133">1.0</span><span class="sxs-lookup"><span data-stu-id="2ed73-133">1.0</span></span>|
|[<span data-ttu-id="2ed73-134">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2ed73-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ed73-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ed73-135">ReadItem</span></span>|
|[<span data-ttu-id="2ed73-136">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2ed73-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ed73-137">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2ed73-137">Compose or Read</span></span>|

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="2ed73-138">hostVersion: строка</span><span class="sxs-lookup"><span data-stu-id="2ed73-138">hostVersion: String</span></span>

<span data-ttu-id="2ed73-139">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="2ed73-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="2ed73-140">Если почтовая надстройка запущена на клиенте Outlook для настольных ПК или iOS `hostVersion` , свойство возвращает версию ведущего приложения, Outlook.</span><span class="sxs-lookup"><span data-stu-id="2ed73-140">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="2ed73-141">В Outlook в Интернете свойство возвращает версию сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="2ed73-141">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="2ed73-142">Пример: строка "15.0.468.0".</span><span class="sxs-lookup"><span data-stu-id="2ed73-142">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="2ed73-143">Тип</span><span class="sxs-lookup"><span data-stu-id="2ed73-143">Type</span></span>

*   <span data-ttu-id="2ed73-144">String</span><span class="sxs-lookup"><span data-stu-id="2ed73-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ed73-145">Требования</span><span class="sxs-lookup"><span data-stu-id="2ed73-145">Requirements</span></span>

|<span data-ttu-id="2ed73-146">Требование</span><span class="sxs-lookup"><span data-stu-id="2ed73-146">Requirement</span></span>| <span data-ttu-id="2ed73-147">Значение</span><span class="sxs-lookup"><span data-stu-id="2ed73-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ed73-148">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2ed73-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ed73-149">1.0</span><span class="sxs-lookup"><span data-stu-id="2ed73-149">1.0</span></span>|
|[<span data-ttu-id="2ed73-150">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2ed73-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ed73-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ed73-151">ReadItem</span></span>|
|[<span data-ttu-id="2ed73-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2ed73-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ed73-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2ed73-153">Compose or Read</span></span>|

---
---

#### <a name="owaview-string"></a><span data-ttu-id="2ed73-154">OWAView: строка</span><span class="sxs-lookup"><span data-stu-id="2ed73-154">OWAView: String</span></span>

<span data-ttu-id="2ed73-155">Получает строку, представляющую текущее представление Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="2ed73-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="2ed73-156">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="2ed73-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="2ed73-157">Если ведущее приложение не является Outlook в Интернете, то при доступе к этому свойству будет получен результат `undefined`.</span><span class="sxs-lookup"><span data-stu-id="2ed73-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="2ed73-158">В Outlook в Интернете есть три представления, которые соответствуют ширине экрана и окна, а также количество отображаемых столбцов:</span><span class="sxs-lookup"><span data-stu-id="2ed73-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="2ed73-159">`OneColumn`, который отображается, когда экран сужается.</span><span class="sxs-lookup"><span data-stu-id="2ed73-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="2ed73-160">В Outlook в Интернете этот макет с одним столбцом используется на всем экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="2ed73-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="2ed73-161">`TwoColumns`, который отображается, когда экран расширяется.</span><span class="sxs-lookup"><span data-stu-id="2ed73-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="2ed73-162">Outlook в Интернете использует это представление на большинстве планшетов.</span><span class="sxs-lookup"><span data-stu-id="2ed73-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="2ed73-163">`ThreeColumns` используется для полноразмерных экранов.</span><span class="sxs-lookup"><span data-stu-id="2ed73-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="2ed73-164">Например, в Outlook в Интернете это представление используется в полноэкранном окне на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="2ed73-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="2ed73-165">Тип</span><span class="sxs-lookup"><span data-stu-id="2ed73-165">Type</span></span>

*   <span data-ttu-id="2ed73-166">String</span><span class="sxs-lookup"><span data-stu-id="2ed73-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2ed73-167">Требования</span><span class="sxs-lookup"><span data-stu-id="2ed73-167">Requirements</span></span>

|<span data-ttu-id="2ed73-168">Требование</span><span class="sxs-lookup"><span data-stu-id="2ed73-168">Requirement</span></span>| <span data-ttu-id="2ed73-169">Значение</span><span class="sxs-lookup"><span data-stu-id="2ed73-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="2ed73-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2ed73-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2ed73-171">1.0</span><span class="sxs-lookup"><span data-stu-id="2ed73-171">1.0</span></span>|
|[<span data-ttu-id="2ed73-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2ed73-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2ed73-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2ed73-173">ReadItem</span></span>|
|[<span data-ttu-id="2ed73-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2ed73-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2ed73-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2ed73-175">Compose or Read</span></span>|
