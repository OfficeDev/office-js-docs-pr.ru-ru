---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,5
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 397a1087d9244589c7b652cc48c6ae0b5ab51d12
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696129"
---
# <a name="diagnostics"></a><span data-ttu-id="c0998-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="c0998-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="c0998-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="c0998-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="c0998-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="c0998-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0998-105">Требования</span><span class="sxs-lookup"><span data-stu-id="c0998-105">Requirements</span></span>

|<span data-ttu-id="c0998-106">Требование</span><span class="sxs-lookup"><span data-stu-id="c0998-106">Requirement</span></span>| <span data-ttu-id="c0998-107">Значение</span><span class="sxs-lookup"><span data-stu-id="c0998-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0998-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c0998-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0998-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c0998-109">1.0</span></span>|
|[<span data-ttu-id="c0998-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c0998-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0998-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0998-111">ReadItem</span></span>|
|[<span data-ttu-id="c0998-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c0998-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0998-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c0998-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c0998-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="c0998-114">Members and methods</span></span>

| <span data-ttu-id="c0998-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="c0998-115">Member</span></span> | <span data-ttu-id="c0998-116">Тип</span><span class="sxs-lookup"><span data-stu-id="c0998-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c0998-117">Сайту</span><span class="sxs-lookup"><span data-stu-id="c0998-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="c0998-118">Member</span><span class="sxs-lookup"><span data-stu-id="c0998-118">Member</span></span> |
| [<span data-ttu-id="c0998-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="c0998-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="c0998-120">Member</span><span class="sxs-lookup"><span data-stu-id="c0998-120">Member</span></span> |
| [<span data-ttu-id="c0998-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="c0998-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="c0998-122">Member</span><span class="sxs-lookup"><span data-stu-id="c0998-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c0998-123">Members</span><span class="sxs-lookup"><span data-stu-id="c0998-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="c0998-124">Имя узла: строка</span><span class="sxs-lookup"><span data-stu-id="c0998-124">hostName: String</span></span>

<span data-ttu-id="c0998-125">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="c0998-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="c0998-126">Строка, которая может иметь одно из следующих значений: `Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="c0998-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

> [!NOTE]
> <span data-ttu-id="c0998-127">`Outlook` Значение возвращается для Outlook на настольных клиентах (например, Windows и Mac).</span><span class="sxs-lookup"><span data-stu-id="c0998-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="c0998-128">Тип</span><span class="sxs-lookup"><span data-stu-id="c0998-128">Type</span></span>

*   <span data-ttu-id="c0998-129">String</span><span class="sxs-lookup"><span data-stu-id="c0998-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0998-130">Требования</span><span class="sxs-lookup"><span data-stu-id="c0998-130">Requirements</span></span>

|<span data-ttu-id="c0998-131">Требование</span><span class="sxs-lookup"><span data-stu-id="c0998-131">Requirement</span></span>| <span data-ttu-id="c0998-132">Значение</span><span class="sxs-lookup"><span data-stu-id="c0998-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0998-133">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c0998-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0998-134">1.0</span><span class="sxs-lookup"><span data-stu-id="c0998-134">1.0</span></span>|
|[<span data-ttu-id="c0998-135">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c0998-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0998-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0998-136">ReadItem</span></span>|
|[<span data-ttu-id="c0998-137">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c0998-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0998-138">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c0998-138">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="c0998-139">hostVersion: строка</span><span class="sxs-lookup"><span data-stu-id="c0998-139">hostVersion: String</span></span>

<span data-ttu-id="c0998-140">Получает строку, представляющую версию ведущего приложения или сервера Exchange (например, "15.0.468.0").</span><span class="sxs-lookup"><span data-stu-id="c0998-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g. "15.0.468.0").</span></span>

<span data-ttu-id="c0998-141">Если почтовая надстройка запущена на клиенте Outlook для настольных ПК или iOS `hostVersion` , свойство возвращает версию ведущего приложения, Outlook.</span><span class="sxs-lookup"><span data-stu-id="c0998-141">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="c0998-142">В Outlook в Интернете свойство возвращает версию сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="c0998-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="c0998-143">Тип</span><span class="sxs-lookup"><span data-stu-id="c0998-143">Type</span></span>

*   <span data-ttu-id="c0998-144">String</span><span class="sxs-lookup"><span data-stu-id="c0998-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0998-145">Требования</span><span class="sxs-lookup"><span data-stu-id="c0998-145">Requirements</span></span>

|<span data-ttu-id="c0998-146">Требование</span><span class="sxs-lookup"><span data-stu-id="c0998-146">Requirement</span></span>| <span data-ttu-id="c0998-147">Значение</span><span class="sxs-lookup"><span data-stu-id="c0998-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0998-148">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c0998-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0998-149">1.0</span><span class="sxs-lookup"><span data-stu-id="c0998-149">1.0</span></span>|
|[<span data-ttu-id="c0998-150">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c0998-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0998-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0998-151">ReadItem</span></span>|
|[<span data-ttu-id="c0998-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c0998-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0998-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c0998-153">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="c0998-154">OWAView: строка</span><span class="sxs-lookup"><span data-stu-id="c0998-154">OWAView: String</span></span>

<span data-ttu-id="c0998-155">Получает строку, представляющую текущее представление Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="c0998-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="c0998-156">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="c0998-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="c0998-157">Если ведущее приложение не является Outlook в Интернете, то при доступе к этому свойству будет получен результат `undefined`.</span><span class="sxs-lookup"><span data-stu-id="c0998-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="c0998-158">В Outlook в Интернете есть три представления, которые соответствуют ширине экрана и окна, а также количество отображаемых столбцов:</span><span class="sxs-lookup"><span data-stu-id="c0998-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="c0998-159">`OneColumn`, который отображается, когда экран сужается.</span><span class="sxs-lookup"><span data-stu-id="c0998-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="c0998-160">В Outlook в Интернете этот макет с одним столбцом используется на всем экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="c0998-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="c0998-161">`TwoColumns`, который отображается, когда экран расширяется.</span><span class="sxs-lookup"><span data-stu-id="c0998-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="c0998-162">Outlook в Интернете использует это представление на большинстве планшетов.</span><span class="sxs-lookup"><span data-stu-id="c0998-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="c0998-163">`ThreeColumns` используется для полноразмерных экранов.</span><span class="sxs-lookup"><span data-stu-id="c0998-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="c0998-164">Например, в Outlook в Интернете это представление используется в полноэкранном окне на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="c0998-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="c0998-165">Тип</span><span class="sxs-lookup"><span data-stu-id="c0998-165">Type</span></span>

*   <span data-ttu-id="c0998-166">String</span><span class="sxs-lookup"><span data-stu-id="c0998-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0998-167">Требования</span><span class="sxs-lookup"><span data-stu-id="c0998-167">Requirements</span></span>

|<span data-ttu-id="c0998-168">Требование</span><span class="sxs-lookup"><span data-stu-id="c0998-168">Requirement</span></span>| <span data-ttu-id="c0998-169">Значение</span><span class="sxs-lookup"><span data-stu-id="c0998-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0998-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c0998-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0998-171">1.0</span><span class="sxs-lookup"><span data-stu-id="c0998-171">1.0</span></span>|
|[<span data-ttu-id="c0998-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c0998-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0998-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0998-173">ReadItem</span></span>|
|[<span data-ttu-id="c0998-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c0998-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c0998-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c0998-175">Compose or Read</span></span>|
