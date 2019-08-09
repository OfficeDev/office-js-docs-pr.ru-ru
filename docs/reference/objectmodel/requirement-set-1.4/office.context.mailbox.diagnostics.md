---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 212afc72cb32bb5c842e1ea606c2edfb64f398ec
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268357"
---
# <a name="diagnostics"></a><span data-ttu-id="f556a-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="f556a-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="f556a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="f556a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="f556a-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="f556a-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f556a-105">Требования</span><span class="sxs-lookup"><span data-stu-id="f556a-105">Requirements</span></span>

|<span data-ttu-id="f556a-106">Требование</span><span class="sxs-lookup"><span data-stu-id="f556a-106">Requirement</span></span>| <span data-ttu-id="f556a-107">Значение</span><span class="sxs-lookup"><span data-stu-id="f556a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f556a-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f556a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f556a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f556a-109">1.0</span></span>|
|[<span data-ttu-id="f556a-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f556a-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f556a-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f556a-111">ReadItem</span></span>|
|[<span data-ttu-id="f556a-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f556a-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f556a-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f556a-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f556a-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="f556a-114">Members and methods</span></span>

| <span data-ttu-id="f556a-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="f556a-115">Member</span></span> | <span data-ttu-id="f556a-116">Тип</span><span class="sxs-lookup"><span data-stu-id="f556a-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f556a-117">Сайту</span><span class="sxs-lookup"><span data-stu-id="f556a-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="f556a-118">Member</span><span class="sxs-lookup"><span data-stu-id="f556a-118">Member</span></span> |
| [<span data-ttu-id="f556a-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="f556a-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="f556a-120">Member</span><span class="sxs-lookup"><span data-stu-id="f556a-120">Member</span></span> |
| [<span data-ttu-id="f556a-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="f556a-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="f556a-122">Member</span><span class="sxs-lookup"><span data-stu-id="f556a-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="f556a-123">Members</span><span class="sxs-lookup"><span data-stu-id="f556a-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="f556a-124">Имя узла: строка</span><span class="sxs-lookup"><span data-stu-id="f556a-124">hostName: String</span></span>

<span data-ttu-id="f556a-125">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="f556a-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="f556a-126">Строка, которая может иметь одно из следующих значений: `Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="f556a-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

> [!NOTE]
> <span data-ttu-id="f556a-127">`Outlook` Значение возвращается для Outlook на настольных клиентах (например, Windows и Mac).</span><span class="sxs-lookup"><span data-stu-id="f556a-127">The `Outlook` value is returned for Outlook on desktop clients (i.e., Windows and Mac).</span></span>

##### <a name="type"></a><span data-ttu-id="f556a-128">Тип</span><span class="sxs-lookup"><span data-stu-id="f556a-128">Type</span></span>

*   <span data-ttu-id="f556a-129">String</span><span class="sxs-lookup"><span data-stu-id="f556a-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f556a-130">Требования</span><span class="sxs-lookup"><span data-stu-id="f556a-130">Requirements</span></span>

|<span data-ttu-id="f556a-131">Требование</span><span class="sxs-lookup"><span data-stu-id="f556a-131">Requirement</span></span>| <span data-ttu-id="f556a-132">Значение</span><span class="sxs-lookup"><span data-stu-id="f556a-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="f556a-133">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f556a-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f556a-134">1.0</span><span class="sxs-lookup"><span data-stu-id="f556a-134">1.0</span></span>|
|[<span data-ttu-id="f556a-135">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f556a-135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f556a-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f556a-136">ReadItem</span></span>|
|[<span data-ttu-id="f556a-137">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f556a-137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f556a-138">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f556a-138">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="f556a-139">hostVersion: строка</span><span class="sxs-lookup"><span data-stu-id="f556a-139">hostVersion: String</span></span>

<span data-ttu-id="f556a-140">Получает строку, представляющую версию ведущего приложения или сервера Exchange (например, "15.0.468.0").</span><span class="sxs-lookup"><span data-stu-id="f556a-140">Gets a string that represents the version of either the host application or the Exchange Server (e.g. "15.0.468.0").</span></span>

<span data-ttu-id="f556a-141">Если почтовая надстройка запущена на клиенте Outlook для настольных ПК или iOS `hostVersion` , свойство возвращает версию ведущего приложения, Outlook.</span><span class="sxs-lookup"><span data-stu-id="f556a-141">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="f556a-142">В Outlook в Интернете свойство возвращает версию сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="f556a-142">In Outlook on the web, the property returns the version of the Exchange Server.</span></span>

##### <a name="type"></a><span data-ttu-id="f556a-143">Тип</span><span class="sxs-lookup"><span data-stu-id="f556a-143">Type</span></span>

*   <span data-ttu-id="f556a-144">String</span><span class="sxs-lookup"><span data-stu-id="f556a-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f556a-145">Требования</span><span class="sxs-lookup"><span data-stu-id="f556a-145">Requirements</span></span>

|<span data-ttu-id="f556a-146">Требование</span><span class="sxs-lookup"><span data-stu-id="f556a-146">Requirement</span></span>| <span data-ttu-id="f556a-147">Значение</span><span class="sxs-lookup"><span data-stu-id="f556a-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="f556a-148">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f556a-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f556a-149">1.0</span><span class="sxs-lookup"><span data-stu-id="f556a-149">1.0</span></span>|
|[<span data-ttu-id="f556a-150">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f556a-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f556a-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f556a-151">ReadItem</span></span>|
|[<span data-ttu-id="f556a-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f556a-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f556a-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f556a-153">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="f556a-154">OWAView: строка</span><span class="sxs-lookup"><span data-stu-id="f556a-154">OWAView: String</span></span>

<span data-ttu-id="f556a-155">Получает строку, представляющую текущее представление Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="f556a-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="f556a-156">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="f556a-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="f556a-157">Если ведущее приложение не является Outlook в Интернете, то при доступе к этому свойству будет получен результат `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f556a-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="f556a-158">В Outlook в Интернете есть три представления, которые соответствуют ширине экрана и окна, а также количество отображаемых столбцов:</span><span class="sxs-lookup"><span data-stu-id="f556a-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="f556a-159">`OneColumn`, который отображается, когда экран сужается.</span><span class="sxs-lookup"><span data-stu-id="f556a-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="f556a-160">В Outlook в Интернете этот макет с одним столбцом используется на всем экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="f556a-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="f556a-161">`TwoColumns`, который отображается, когда экран расширяется.</span><span class="sxs-lookup"><span data-stu-id="f556a-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="f556a-162">Outlook в Интернете использует это представление на большинстве планшетов.</span><span class="sxs-lookup"><span data-stu-id="f556a-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="f556a-163">`ThreeColumns` используется для полноразмерных экранов.</span><span class="sxs-lookup"><span data-stu-id="f556a-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="f556a-164">Например, в Outlook в Интернете это представление используется в полноэкранном окне на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="f556a-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="f556a-165">Тип</span><span class="sxs-lookup"><span data-stu-id="f556a-165">Type</span></span>

*   <span data-ttu-id="f556a-166">String</span><span class="sxs-lookup"><span data-stu-id="f556a-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f556a-167">Требования</span><span class="sxs-lookup"><span data-stu-id="f556a-167">Requirements</span></span>

|<span data-ttu-id="f556a-168">Требование</span><span class="sxs-lookup"><span data-stu-id="f556a-168">Requirement</span></span>| <span data-ttu-id="f556a-169">Значение</span><span class="sxs-lookup"><span data-stu-id="f556a-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="f556a-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f556a-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f556a-171">1.0</span><span class="sxs-lookup"><span data-stu-id="f556a-171">1.0</span></span>|
|[<span data-ttu-id="f556a-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f556a-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f556a-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f556a-173">ReadItem</span></span>|
|[<span data-ttu-id="f556a-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f556a-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f556a-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f556a-175">Compose or Read</span></span>|
