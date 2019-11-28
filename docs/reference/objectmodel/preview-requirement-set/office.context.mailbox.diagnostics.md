---
title: Office. Context. Mailbox. Diagnostics — Предварительная версия набора требований
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 492e292737417854adfaf98feb2b67788933d874
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629204"
---
# <a name="diagnostics"></a><span data-ttu-id="93664-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="93664-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="93664-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="93664-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="93664-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="93664-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="93664-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="93664-105">Requirements</span></span>

|<span data-ttu-id="93664-106">Требование</span><span class="sxs-lookup"><span data-stu-id="93664-106">Requirement</span></span>| <span data-ttu-id="93664-107">Значение</span><span class="sxs-lookup"><span data-stu-id="93664-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="93664-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="93664-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93664-109">1.0</span><span class="sxs-lookup"><span data-stu-id="93664-109">1.0</span></span>|
|[<span data-ttu-id="93664-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="93664-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="93664-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="93664-111">ReadItem</span></span>|
|[<span data-ttu-id="93664-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="93664-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="93664-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="93664-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="93664-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="93664-114">Properties</span></span>

| <span data-ttu-id="93664-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="93664-115">Property</span></span> | <span data-ttu-id="93664-116">Минимальные</span><span class="sxs-lookup"><span data-stu-id="93664-116">Minimum</span></span><br><span data-ttu-id="93664-117">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="93664-117">permission level</span></span> | <span data-ttu-id="93664-118">Способов</span><span class="sxs-lookup"><span data-stu-id="93664-118">Modes</span></span> | <span data-ttu-id="93664-119">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="93664-119">Return type</span></span> | <span data-ttu-id="93664-120">Минимальные</span><span class="sxs-lookup"><span data-stu-id="93664-120">Minimum</span></span><br><span data-ttu-id="93664-121">набор требований</span><span class="sxs-lookup"><span data-stu-id="93664-121">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="93664-122">Сайту</span><span class="sxs-lookup"><span data-stu-id="93664-122">hostName</span></span>](#hostname-string) | <span data-ttu-id="93664-123">ReadItem</span><span class="sxs-lookup"><span data-stu-id="93664-123">ReadItem</span></span> | <span data-ttu-id="93664-124">Создание</span><span class="sxs-lookup"><span data-stu-id="93664-124">Compose</span></span><br><span data-ttu-id="93664-125">Чтение</span><span class="sxs-lookup"><span data-stu-id="93664-125">Read</span></span> | <span data-ttu-id="93664-126">String</span><span class="sxs-lookup"><span data-stu-id="93664-126">String</span></span> | <span data-ttu-id="93664-127">1.0</span><span class="sxs-lookup"><span data-stu-id="93664-127">1.0</span></span> |
| [<span data-ttu-id="93664-128">hostVersion</span><span class="sxs-lookup"><span data-stu-id="93664-128">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="93664-129">ReadItem</span><span class="sxs-lookup"><span data-stu-id="93664-129">ReadItem</span></span> | <span data-ttu-id="93664-130">Создание</span><span class="sxs-lookup"><span data-stu-id="93664-130">Compose</span></span><br><span data-ttu-id="93664-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="93664-131">Read</span></span> | <span data-ttu-id="93664-132">String</span><span class="sxs-lookup"><span data-stu-id="93664-132">String</span></span> | <span data-ttu-id="93664-133">1.0</span><span class="sxs-lookup"><span data-stu-id="93664-133">1.0</span></span> |
| [<span data-ttu-id="93664-134">OWAView</span><span class="sxs-lookup"><span data-stu-id="93664-134">OWAView</span></span>](#owaview-string) | <span data-ttu-id="93664-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="93664-135">ReadItem</span></span> | <span data-ttu-id="93664-136">Создание</span><span class="sxs-lookup"><span data-stu-id="93664-136">Compose</span></span><br><span data-ttu-id="93664-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="93664-137">Read</span></span> | <span data-ttu-id="93664-138">String</span><span class="sxs-lookup"><span data-stu-id="93664-138">String</span></span> | <span data-ttu-id="93664-139">1.0</span><span class="sxs-lookup"><span data-stu-id="93664-139">1.0</span></span> |

## <a name="property-details"></a><span data-ttu-id="93664-140">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="93664-140">Property details</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="93664-141">Имя узла: строка</span><span class="sxs-lookup"><span data-stu-id="93664-141">hostName: String</span></span>

<span data-ttu-id="93664-142">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="93664-142">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="93664-143">Строка, которая может иметь одно из следующих значений: `Outlook`, `OutlookWebApp`, `OutlookIOS` или `OutlookAndroid`.</span><span class="sxs-lookup"><span data-stu-id="93664-143">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

> [!NOTE]
> <span data-ttu-id="93664-144">`Outlook` Значение возвращается для Outlook на настольных клиентах (например, Windows и Mac).</span><span class="sxs-lookup"><span data-stu-id="93664-144">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="93664-145">Тип</span><span class="sxs-lookup"><span data-stu-id="93664-145">Type</span></span>

*   <span data-ttu-id="93664-146">String</span><span class="sxs-lookup"><span data-stu-id="93664-146">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="93664-147">Требования</span><span class="sxs-lookup"><span data-stu-id="93664-147">Requirements</span></span>

|<span data-ttu-id="93664-148">Требование</span><span class="sxs-lookup"><span data-stu-id="93664-148">Requirement</span></span>| <span data-ttu-id="93664-149">Значение</span><span class="sxs-lookup"><span data-stu-id="93664-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="93664-150">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="93664-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93664-151">1.0</span><span class="sxs-lookup"><span data-stu-id="93664-151">1.0</span></span>|
|[<span data-ttu-id="93664-152">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="93664-152">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="93664-153">ReadItem</span><span class="sxs-lookup"><span data-stu-id="93664-153">ReadItem</span></span>|
|[<span data-ttu-id="93664-154">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="93664-154">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="93664-155">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="93664-155">Compose or Read</span></span>|

<br>

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="93664-156">hostVersion: строка</span><span class="sxs-lookup"><span data-stu-id="93664-156">hostVersion: String</span></span>

<span data-ttu-id="93664-157">Получает строку, представляющую версию ведущего приложения или сервера Exchange (например, "15.0.468.0").</span><span class="sxs-lookup"><span data-stu-id="93664-157">Gets a string that represents the version of either the host application or the Exchange Server (e.g., "15.0.468.0").</span></span>

<span data-ttu-id="93664-158">Если почтовая надстройка запущена на настольном клиенте Outlook или мобильном клиенте, `hostVersion` свойство возвращает версию ведущего приложения, Outlook.</span><span class="sxs-lookup"><span data-stu-id="93664-158">If the mail add-in is running on an Outlook desktop or mobile client, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="93664-159">В Outlook в Интернете свойство возвращает версию сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="93664-159">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="93664-160">Тип</span><span class="sxs-lookup"><span data-stu-id="93664-160">Type</span></span>

*   <span data-ttu-id="93664-161">String</span><span class="sxs-lookup"><span data-stu-id="93664-161">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="93664-162">Требования</span><span class="sxs-lookup"><span data-stu-id="93664-162">Requirements</span></span>

|<span data-ttu-id="93664-163">Требование</span><span class="sxs-lookup"><span data-stu-id="93664-163">Requirement</span></span>| <span data-ttu-id="93664-164">Значение</span><span class="sxs-lookup"><span data-stu-id="93664-164">Value</span></span>|
|---|---|
|[<span data-ttu-id="93664-165">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="93664-165">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93664-166">1.0</span><span class="sxs-lookup"><span data-stu-id="93664-166">1.0</span></span>|
|[<span data-ttu-id="93664-167">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="93664-167">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="93664-168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="93664-168">ReadItem</span></span>|
|[<span data-ttu-id="93664-169">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="93664-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="93664-170">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="93664-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="owaview-string"></a><span data-ttu-id="93664-171">OWAView: строка</span><span class="sxs-lookup"><span data-stu-id="93664-171">OWAView: String</span></span>

<span data-ttu-id="93664-172">Получает строку, представляющую текущее представление Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="93664-172">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="93664-173">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="93664-173">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="93664-174">Если ведущее приложение не является Outlook в Интернете, то при доступе к этому свойству будет получен результат `undefined`.</span><span class="sxs-lookup"><span data-stu-id="93664-174">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="93664-175">В Outlook в Интернете есть три представления, которые соответствуют ширине экрана и окна, а также количество отображаемых столбцов:</span><span class="sxs-lookup"><span data-stu-id="93664-175">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="93664-176">`OneColumn`, который отображается, когда экран сужается.</span><span class="sxs-lookup"><span data-stu-id="93664-176">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="93664-177">В Outlook в Интернете этот макет с одним столбцом используется на всем экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="93664-177">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="93664-178">`TwoColumns`, который отображается, когда экран расширяется.</span><span class="sxs-lookup"><span data-stu-id="93664-178">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="93664-179">Outlook в Интернете использует это представление на большинстве планшетов.</span><span class="sxs-lookup"><span data-stu-id="93664-179">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="93664-180">`ThreeColumns` используется для полноразмерных экранов.</span><span class="sxs-lookup"><span data-stu-id="93664-180">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="93664-181">Например, в Outlook в Интернете это представление используется в полноэкранном окне на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="93664-181">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="93664-182">Тип</span><span class="sxs-lookup"><span data-stu-id="93664-182">Type</span></span>

*   <span data-ttu-id="93664-183">String</span><span class="sxs-lookup"><span data-stu-id="93664-183">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="93664-184">Требования</span><span class="sxs-lookup"><span data-stu-id="93664-184">Requirements</span></span>

|<span data-ttu-id="93664-185">Требование</span><span class="sxs-lookup"><span data-stu-id="93664-185">Requirement</span></span>| <span data-ttu-id="93664-186">Значение</span><span class="sxs-lookup"><span data-stu-id="93664-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="93664-187">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="93664-187">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="93664-188">1.0</span><span class="sxs-lookup"><span data-stu-id="93664-188">1.0</span></span>|
|[<span data-ttu-id="93664-189">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="93664-189">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="93664-190">ReadItem</span><span class="sxs-lookup"><span data-stu-id="93664-190">ReadItem</span></span>|
|[<span data-ttu-id="93664-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="93664-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="93664-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="93664-192">Compose or Read</span></span>|
