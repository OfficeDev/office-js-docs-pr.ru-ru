---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,8
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 8b2d67fbc5eb8462af67a0dc73ce65a433ad5795
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902203"
---
# <a name="diagnostics"></a><span data-ttu-id="65f4b-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="65f4b-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="65f4b-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="65f4b-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="65f4b-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="65f4b-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="65f4b-105">Требования</span><span class="sxs-lookup"><span data-stu-id="65f4b-105">Requirements</span></span>

|<span data-ttu-id="65f4b-106">Требование</span><span class="sxs-lookup"><span data-stu-id="65f4b-106">Requirement</span></span>| <span data-ttu-id="65f4b-107">Значение</span><span class="sxs-lookup"><span data-stu-id="65f4b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="65f4b-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="65f4b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="65f4b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="65f4b-109">1.0</span></span>|
|[<span data-ttu-id="65f4b-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="65f4b-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="65f4b-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="65f4b-111">ReadItem</span></span>|
|[<span data-ttu-id="65f4b-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="65f4b-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="65f4b-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="65f4b-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="65f4b-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="65f4b-114">Members and methods</span></span>

| <span data-ttu-id="65f4b-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="65f4b-115">Member</span></span> | <span data-ttu-id="65f4b-116">Тип</span><span class="sxs-lookup"><span data-stu-id="65f4b-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="65f4b-117">Сайту</span><span class="sxs-lookup"><span data-stu-id="65f4b-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="65f4b-118">Member</span><span class="sxs-lookup"><span data-stu-id="65f4b-118">Member</span></span> |
| [<span data-ttu-id="65f4b-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="65f4b-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="65f4b-120">Member</span><span class="sxs-lookup"><span data-stu-id="65f4b-120">Member</span></span> |
| [<span data-ttu-id="65f4b-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="65f4b-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="65f4b-122">Member</span><span class="sxs-lookup"><span data-stu-id="65f4b-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="65f4b-123">Members</span><span class="sxs-lookup"><span data-stu-id="65f4b-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="65f4b-124">Имя узла: строка</span><span class="sxs-lookup"><span data-stu-id="65f4b-124">hostName: String</span></span>

<span data-ttu-id="65f4b-125">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="65f4b-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="65f4b-126">Строка, которая может иметь одно из следующих значений: `Outlook`, `OutlookWebApp`, `OutlookIOS` или `OutlookAndroid`.</span><span class="sxs-lookup"><span data-stu-id="65f4b-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="65f4b-127">`Outlook` Значение возвращается для Outlook на настольных клиентах (например, Windows и Mac).</span><span class="sxs-lookup"><span data-stu-id="65f4b-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="65f4b-128">Тип</span><span class="sxs-lookup"><span data-stu-id="65f4b-128">Type</span></span>

*   <span data-ttu-id="65f4b-129">String</span><span class="sxs-lookup"><span data-stu-id="65f4b-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="65f4b-130">Требования</span><span class="sxs-lookup"><span data-stu-id="65f4b-130">Requirements</span></span>

|<span data-ttu-id="65f4b-131">Требование</span><span class="sxs-lookup"><span data-stu-id="65f4b-131">Requirement</span></span>| <span data-ttu-id="65f4b-132">Значение</span><span class="sxs-lookup"><span data-stu-id="65f4b-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="65f4b-133">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="65f4b-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="65f4b-134">1.0</span><span class="sxs-lookup"><span data-stu-id="65f4b-134">1.0</span></span>|
|[<span data-ttu-id="65f4b-135">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="65f4b-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="65f4b-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="65f4b-136">ReadItem</span></span>|
|[<span data-ttu-id="65f4b-137">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="65f4b-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="65f4b-138">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="65f4b-138">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="65f4b-139">hostVersion: строка</span><span class="sxs-lookup"><span data-stu-id="65f4b-139">hostVersion: String</span></span>

<span data-ttu-id="65f4b-140">Получает строку, представляющую версию ведущего приложения или сервера Exchange (например, "15.0.468.0").</span><span class="sxs-lookup"><span data-stu-id="65f4b-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="65f4b-141">Если почтовая надстройка работает в клиенте Outlook для настольного компьютера или в iOS `hostVersion` , свойство возвращает версию ведущего приложения, Outlook.</span><span class="sxs-lookup"><span data-stu-id="65f4b-141">If the mail add-in is running on the Outlook desktop client or on iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="65f4b-142">В Outlook в Интернете свойство возвращает версию сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="65f4b-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="65f4b-143">Тип</span><span class="sxs-lookup"><span data-stu-id="65f4b-143">Type</span></span>

*   <span data-ttu-id="65f4b-144">String</span><span class="sxs-lookup"><span data-stu-id="65f4b-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="65f4b-145">Требования</span><span class="sxs-lookup"><span data-stu-id="65f4b-145">Requirements</span></span>

|<span data-ttu-id="65f4b-146">Требование</span><span class="sxs-lookup"><span data-stu-id="65f4b-146">Requirement</span></span>| <span data-ttu-id="65f4b-147">Значение</span><span class="sxs-lookup"><span data-stu-id="65f4b-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="65f4b-148">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="65f4b-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="65f4b-149">1.0</span><span class="sxs-lookup"><span data-stu-id="65f4b-149">1.0</span></span>|
|[<span data-ttu-id="65f4b-150">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="65f4b-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="65f4b-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="65f4b-151">ReadItem</span></span>|
|[<span data-ttu-id="65f4b-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="65f4b-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="65f4b-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="65f4b-153">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="65f4b-154">OWAView: строка</span><span class="sxs-lookup"><span data-stu-id="65f4b-154">OWAView: String</span></span>

<span data-ttu-id="65f4b-155">Получает строку, представляющую текущее представление Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="65f4b-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="65f4b-156">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="65f4b-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="65f4b-157">Если ведущее приложение не является Outlook в Интернете, то при доступе к этому свойству будет получен результат `undefined`.</span><span class="sxs-lookup"><span data-stu-id="65f4b-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="65f4b-158">В Outlook в Интернете есть три представления, которые соответствуют ширине экрана и окна, а также количество отображаемых столбцов:</span><span class="sxs-lookup"><span data-stu-id="65f4b-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="65f4b-159">`OneColumn`, который отображается, когда экран сужается.</span><span class="sxs-lookup"><span data-stu-id="65f4b-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="65f4b-160">В Outlook в Интернете этот макет с одним столбцом используется на всем экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="65f4b-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="65f4b-161">`TwoColumns`, который отображается, когда экран расширяется.</span><span class="sxs-lookup"><span data-stu-id="65f4b-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="65f4b-162">Outlook в Интернете использует это представление на большинстве планшетов.</span><span class="sxs-lookup"><span data-stu-id="65f4b-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="65f4b-163">`ThreeColumns` используется для полноразмерных экранов.</span><span class="sxs-lookup"><span data-stu-id="65f4b-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="65f4b-164">Например, в Outlook в Интернете это представление используется в полноэкранном окне на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="65f4b-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="65f4b-165">Тип</span><span class="sxs-lookup"><span data-stu-id="65f4b-165">Type</span></span>

*   <span data-ttu-id="65f4b-166">String</span><span class="sxs-lookup"><span data-stu-id="65f4b-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="65f4b-167">Требования</span><span class="sxs-lookup"><span data-stu-id="65f4b-167">Requirements</span></span>

|<span data-ttu-id="65f4b-168">Требование</span><span class="sxs-lookup"><span data-stu-id="65f4b-168">Requirement</span></span>| <span data-ttu-id="65f4b-169">Значение</span><span class="sxs-lookup"><span data-stu-id="65f4b-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="65f4b-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="65f4b-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="65f4b-171">1.0</span><span class="sxs-lookup"><span data-stu-id="65f4b-171">1.0</span></span>|
|[<span data-ttu-id="65f4b-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="65f4b-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="65f4b-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="65f4b-173">ReadItem</span></span>|
|[<span data-ttu-id="65f4b-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="65f4b-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="65f4b-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="65f4b-175">Compose or Read</span></span>|
