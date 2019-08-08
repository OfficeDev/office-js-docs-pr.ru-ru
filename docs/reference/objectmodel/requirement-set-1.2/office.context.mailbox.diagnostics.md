---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,2
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: f2e613816884a5c1c00e5b96565d378434747e8e
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231272"
---
# <a name="diagnostics"></a><span data-ttu-id="43170-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="43170-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="43170-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="43170-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="43170-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="43170-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="43170-105">Требования</span><span class="sxs-lookup"><span data-stu-id="43170-105">Requirements</span></span>

|<span data-ttu-id="43170-106">Требование</span><span class="sxs-lookup"><span data-stu-id="43170-106">Requirement</span></span>| <span data-ttu-id="43170-107">Значение</span><span class="sxs-lookup"><span data-stu-id="43170-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="43170-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="43170-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43170-109">1.0</span><span class="sxs-lookup"><span data-stu-id="43170-109">1.0</span></span>|
|[<span data-ttu-id="43170-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="43170-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43170-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43170-111">ReadItem</span></span>|
|[<span data-ttu-id="43170-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="43170-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43170-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="43170-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="43170-114">Members</span><span class="sxs-lookup"><span data-stu-id="43170-114">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="43170-115">Имя узла: строка</span><span class="sxs-lookup"><span data-stu-id="43170-115">hostName: String</span></span>

<span data-ttu-id="43170-116">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="43170-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="43170-117">Строка, которая может иметь одно из следующих значений: `Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="43170-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="43170-118">Тип</span><span class="sxs-lookup"><span data-stu-id="43170-118">Type</span></span>

*   <span data-ttu-id="43170-119">String</span><span class="sxs-lookup"><span data-stu-id="43170-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43170-120">Требования</span><span class="sxs-lookup"><span data-stu-id="43170-120">Requirements</span></span>

|<span data-ttu-id="43170-121">Требование</span><span class="sxs-lookup"><span data-stu-id="43170-121">Requirement</span></span>| <span data-ttu-id="43170-122">Значение</span><span class="sxs-lookup"><span data-stu-id="43170-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="43170-123">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="43170-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43170-124">1.0</span><span class="sxs-lookup"><span data-stu-id="43170-124">1.0</span></span>|
|[<span data-ttu-id="43170-125">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="43170-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43170-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43170-126">ReadItem</span></span>|
|[<span data-ttu-id="43170-127">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="43170-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43170-128">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="43170-128">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="43170-129">hostVersion: строка</span><span class="sxs-lookup"><span data-stu-id="43170-129">hostVersion: String</span></span>

<span data-ttu-id="43170-130">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="43170-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="43170-131">Если почтовая надстройка запущена на клиенте Outlook для настольных ПК или iOS `hostVersion` , свойство возвращает версию ведущего приложения, Outlook.</span><span class="sxs-lookup"><span data-stu-id="43170-131">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="43170-132">В Outlook в Интернете свойство возвращает версию сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="43170-132">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="43170-133">Пример: строка "15.0.468.0".</span><span class="sxs-lookup"><span data-stu-id="43170-133">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="43170-134">Тип</span><span class="sxs-lookup"><span data-stu-id="43170-134">Type</span></span>

*   <span data-ttu-id="43170-135">String</span><span class="sxs-lookup"><span data-stu-id="43170-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43170-136">Требования</span><span class="sxs-lookup"><span data-stu-id="43170-136">Requirements</span></span>

|<span data-ttu-id="43170-137">Требование</span><span class="sxs-lookup"><span data-stu-id="43170-137">Requirement</span></span>| <span data-ttu-id="43170-138">Значение</span><span class="sxs-lookup"><span data-stu-id="43170-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="43170-139">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="43170-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43170-140">1.0</span><span class="sxs-lookup"><span data-stu-id="43170-140">1.0</span></span>|
|[<span data-ttu-id="43170-141">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="43170-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43170-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43170-142">ReadItem</span></span>|
|[<span data-ttu-id="43170-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="43170-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43170-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="43170-144">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="43170-145">OWAView: строка</span><span class="sxs-lookup"><span data-stu-id="43170-145">OWAView: String</span></span>

<span data-ttu-id="43170-146">Получает строку, представляющую текущее представление Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="43170-146">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="43170-147">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="43170-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="43170-148">Если ведущее приложение не является Outlook в Интернете, то при доступе к этому свойству будет получен результат `undefined`.</span><span class="sxs-lookup"><span data-stu-id="43170-148">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="43170-149">В Outlook в Интернете есть три представления, которые соответствуют ширине экрана и окна, а также количество отображаемых столбцов:</span><span class="sxs-lookup"><span data-stu-id="43170-149">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="43170-150">`OneColumn`, который отображается, когда экран сужается.</span><span class="sxs-lookup"><span data-stu-id="43170-150">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="43170-151">В Outlook в Интернете этот макет с одним столбцом используется на всем экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="43170-151">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="43170-152">`TwoColumns`, который отображается, когда экран расширяется.</span><span class="sxs-lookup"><span data-stu-id="43170-152">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="43170-153">Outlook в Интернете использует это представление на большинстве планшетов.</span><span class="sxs-lookup"><span data-stu-id="43170-153">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="43170-154">`ThreeColumns` используется для полноразмерных экранов.</span><span class="sxs-lookup"><span data-stu-id="43170-154">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="43170-155">Например, в Outlook в Интернете это представление используется в полноэкранном окне на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="43170-155">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="43170-156">Тип</span><span class="sxs-lookup"><span data-stu-id="43170-156">Type</span></span>

*   <span data-ttu-id="43170-157">String</span><span class="sxs-lookup"><span data-stu-id="43170-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="43170-158">Требования</span><span class="sxs-lookup"><span data-stu-id="43170-158">Requirements</span></span>

|<span data-ttu-id="43170-159">Требование</span><span class="sxs-lookup"><span data-stu-id="43170-159">Requirement</span></span>| <span data-ttu-id="43170-160">Значение</span><span class="sxs-lookup"><span data-stu-id="43170-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="43170-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="43170-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="43170-162">1.0</span><span class="sxs-lookup"><span data-stu-id="43170-162">1.0</span></span>|
|[<span data-ttu-id="43170-163">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="43170-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="43170-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="43170-164">ReadItem</span></span>|
|[<span data-ttu-id="43170-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="43170-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="43170-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="43170-166">Compose or Read</span></span>|
