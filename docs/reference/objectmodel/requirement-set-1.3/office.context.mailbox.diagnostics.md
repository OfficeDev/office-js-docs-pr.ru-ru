---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 696aa8769b52b0f96d4a68292c156394ed6be2a2
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268679"
---
# <a name="diagnostics"></a><span data-ttu-id="75262-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="75262-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="75262-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="75262-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="75262-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="75262-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="75262-105">Требования</span><span class="sxs-lookup"><span data-stu-id="75262-105">Requirements</span></span>

|<span data-ttu-id="75262-106">Требование</span><span class="sxs-lookup"><span data-stu-id="75262-106">Requirement</span></span>| <span data-ttu-id="75262-107">Значение</span><span class="sxs-lookup"><span data-stu-id="75262-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="75262-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="75262-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75262-109">1.0</span><span class="sxs-lookup"><span data-stu-id="75262-109">1.0</span></span>|
|[<span data-ttu-id="75262-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="75262-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75262-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75262-111">ReadItem</span></span>|
|[<span data-ttu-id="75262-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="75262-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="75262-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="75262-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="75262-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="75262-114">Members and methods</span></span>

| <span data-ttu-id="75262-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="75262-115">Member</span></span> | <span data-ttu-id="75262-116">Тип</span><span class="sxs-lookup"><span data-stu-id="75262-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="75262-117">Сайту</span><span class="sxs-lookup"><span data-stu-id="75262-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="75262-118">Member</span><span class="sxs-lookup"><span data-stu-id="75262-118">Member</span></span> |
| [<span data-ttu-id="75262-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="75262-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="75262-120">Member</span><span class="sxs-lookup"><span data-stu-id="75262-120">Member</span></span> |
| [<span data-ttu-id="75262-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="75262-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="75262-122">Member</span><span class="sxs-lookup"><span data-stu-id="75262-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="75262-123">Members</span><span class="sxs-lookup"><span data-stu-id="75262-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="75262-124">Имя узла: строка</span><span class="sxs-lookup"><span data-stu-id="75262-124">hostName: String</span></span>

<span data-ttu-id="75262-125">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="75262-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="75262-126">Строка, которая может иметь одно из следующих значений: `Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="75262-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

> [!NOTE]
> <span data-ttu-id="75262-127">`Outlook` Значение возвращается для Outlook на настольных клиентах (например, Windows и Mac).</span><span class="sxs-lookup"><span data-stu-id="75262-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="75262-128">Тип</span><span class="sxs-lookup"><span data-stu-id="75262-128">Type</span></span>

*   <span data-ttu-id="75262-129">String</span><span class="sxs-lookup"><span data-stu-id="75262-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75262-130">Требования</span><span class="sxs-lookup"><span data-stu-id="75262-130">Requirements</span></span>

|<span data-ttu-id="75262-131">Требование</span><span class="sxs-lookup"><span data-stu-id="75262-131">Requirement</span></span>| <span data-ttu-id="75262-132">Значение</span><span class="sxs-lookup"><span data-stu-id="75262-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="75262-133">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="75262-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75262-134">1.0</span><span class="sxs-lookup"><span data-stu-id="75262-134">1.0</span></span>|
|[<span data-ttu-id="75262-135">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="75262-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75262-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75262-136">ReadItem</span></span>|
|[<span data-ttu-id="75262-137">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="75262-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="75262-138">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="75262-138">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="75262-139">hostVersion: строка</span><span class="sxs-lookup"><span data-stu-id="75262-139">hostVersion: String</span></span>

<span data-ttu-id="75262-140">Получает строку, представляющую версию ведущего приложения или сервера Exchange (например, "15.0.468.0").</span><span class="sxs-lookup"><span data-stu-id="75262-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="75262-141">Если почтовая надстройка запущена на клиенте Outlook для настольных ПК или iOS `hostVersion` , свойство возвращает версию ведущего приложения, Outlook.</span><span class="sxs-lookup"><span data-stu-id="75262-141">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="75262-142">В Outlook в Интернете свойство возвращает версию сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="75262-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="75262-143">Тип</span><span class="sxs-lookup"><span data-stu-id="75262-143">Type</span></span>

*   <span data-ttu-id="75262-144">String</span><span class="sxs-lookup"><span data-stu-id="75262-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75262-145">Требования</span><span class="sxs-lookup"><span data-stu-id="75262-145">Requirements</span></span>

|<span data-ttu-id="75262-146">Требование</span><span class="sxs-lookup"><span data-stu-id="75262-146">Requirement</span></span>| <span data-ttu-id="75262-147">Значение</span><span class="sxs-lookup"><span data-stu-id="75262-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="75262-148">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="75262-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75262-149">1.0</span><span class="sxs-lookup"><span data-stu-id="75262-149">1.0</span></span>|
|[<span data-ttu-id="75262-150">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="75262-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75262-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75262-151">ReadItem</span></span>|
|[<span data-ttu-id="75262-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="75262-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="75262-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="75262-153">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="75262-154">OWAView: строка</span><span class="sxs-lookup"><span data-stu-id="75262-154">OWAView: String</span></span>

<span data-ttu-id="75262-155">Получает строку, представляющую текущее представление Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="75262-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="75262-156">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="75262-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="75262-157">Если ведущее приложение не является Outlook в Интернете, то при доступе к этому свойству будет получен результат `undefined`.</span><span class="sxs-lookup"><span data-stu-id="75262-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="75262-158">В Outlook в Интернете есть три представления, которые соответствуют ширине экрана и окна, а также количество отображаемых столбцов:</span><span class="sxs-lookup"><span data-stu-id="75262-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="75262-159">`OneColumn`, который отображается, когда экран сужается.</span><span class="sxs-lookup"><span data-stu-id="75262-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="75262-160">В Outlook в Интернете этот макет с одним столбцом используется на всем экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="75262-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="75262-161">`TwoColumns`, который отображается, когда экран расширяется.</span><span class="sxs-lookup"><span data-stu-id="75262-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="75262-162">Outlook в Интернете использует это представление на большинстве планшетов.</span><span class="sxs-lookup"><span data-stu-id="75262-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="75262-163">`ThreeColumns` используется для полноразмерных экранов.</span><span class="sxs-lookup"><span data-stu-id="75262-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="75262-164">Например, в Outlook в Интернете это представление используется в полноэкранном окне на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="75262-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="75262-165">Тип</span><span class="sxs-lookup"><span data-stu-id="75262-165">Type</span></span>

*   <span data-ttu-id="75262-166">String</span><span class="sxs-lookup"><span data-stu-id="75262-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="75262-167">Требования</span><span class="sxs-lookup"><span data-stu-id="75262-167">Requirements</span></span>

|<span data-ttu-id="75262-168">Требование</span><span class="sxs-lookup"><span data-stu-id="75262-168">Requirement</span></span>| <span data-ttu-id="75262-169">Значение</span><span class="sxs-lookup"><span data-stu-id="75262-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="75262-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="75262-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="75262-171">1.0</span><span class="sxs-lookup"><span data-stu-id="75262-171">1.0</span></span>|
|[<span data-ttu-id="75262-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="75262-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="75262-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="75262-173">ReadItem</span></span>|
|[<span data-ttu-id="75262-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="75262-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="75262-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="75262-175">Compose or Read</span></span>|
