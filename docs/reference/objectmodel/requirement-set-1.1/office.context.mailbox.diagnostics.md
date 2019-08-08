---
title: Office.context.mailbox.diagnostics — набор обязательных элементов 1.1
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 32c4d7389555bf6dab7cff3822abb4a4edbebe80
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231300"
---
# <a name="diagnostics"></a><span data-ttu-id="f348b-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="f348b-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="f348b-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="f348b-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="f348b-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="f348b-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f348b-105">Требования</span><span class="sxs-lookup"><span data-stu-id="f348b-105">Requirements</span></span>

|<span data-ttu-id="f348b-106">Требование</span><span class="sxs-lookup"><span data-stu-id="f348b-106">Requirement</span></span>| <span data-ttu-id="f348b-107">Значение</span><span class="sxs-lookup"><span data-stu-id="f348b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f348b-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f348b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f348b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f348b-109">1.0</span></span>|
|[<span data-ttu-id="f348b-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f348b-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f348b-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f348b-111">ReadItem</span></span>|
|[<span data-ttu-id="f348b-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f348b-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f348b-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f348b-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="f348b-114">Members</span><span class="sxs-lookup"><span data-stu-id="f348b-114">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="f348b-115">Имя узла: строка</span><span class="sxs-lookup"><span data-stu-id="f348b-115">hostName: String</span></span>

<span data-ttu-id="f348b-116">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="f348b-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="f348b-117">Строка, которая может иметь одно из следующих значений: `Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="f348b-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="f348b-118">Тип</span><span class="sxs-lookup"><span data-stu-id="f348b-118">Type</span></span>

*   <span data-ttu-id="f348b-119">String</span><span class="sxs-lookup"><span data-stu-id="f348b-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f348b-120">Требования</span><span class="sxs-lookup"><span data-stu-id="f348b-120">Requirements</span></span>

|<span data-ttu-id="f348b-121">Требование</span><span class="sxs-lookup"><span data-stu-id="f348b-121">Requirement</span></span>| <span data-ttu-id="f348b-122">Значение</span><span class="sxs-lookup"><span data-stu-id="f348b-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="f348b-123">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f348b-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f348b-124">1.0</span><span class="sxs-lookup"><span data-stu-id="f348b-124">1.0</span></span>|
|[<span data-ttu-id="f348b-125">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f348b-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f348b-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f348b-126">ReadItem</span></span>|
|[<span data-ttu-id="f348b-127">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f348b-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f348b-128">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f348b-128">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="f348b-129">hostVersion: строка</span><span class="sxs-lookup"><span data-stu-id="f348b-129">hostVersion: String</span></span>

<span data-ttu-id="f348b-130">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="f348b-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="f348b-131">Если почтовая надстройка запущена на клиенте Outlook для настольных ПК или iOS `hostVersion` , свойство возвращает версию ведущего приложения, Outlook.</span><span class="sxs-lookup"><span data-stu-id="f348b-131">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="f348b-132">В Outlook в Интернете свойство возвращает версию сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="f348b-132">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="f348b-133">Пример: строка "15.0.468.0".</span><span class="sxs-lookup"><span data-stu-id="f348b-133">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="f348b-134">Тип</span><span class="sxs-lookup"><span data-stu-id="f348b-134">Type</span></span>

*   <span data-ttu-id="f348b-135">String</span><span class="sxs-lookup"><span data-stu-id="f348b-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f348b-136">Требования</span><span class="sxs-lookup"><span data-stu-id="f348b-136">Requirements</span></span>

|<span data-ttu-id="f348b-137">Требование</span><span class="sxs-lookup"><span data-stu-id="f348b-137">Requirement</span></span>| <span data-ttu-id="f348b-138">Значение</span><span class="sxs-lookup"><span data-stu-id="f348b-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="f348b-139">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f348b-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f348b-140">1.0</span><span class="sxs-lookup"><span data-stu-id="f348b-140">1.0</span></span>|
|[<span data-ttu-id="f348b-141">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f348b-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f348b-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f348b-142">ReadItem</span></span>|
|[<span data-ttu-id="f348b-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f348b-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f348b-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f348b-144">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="f348b-145">OWAView: строка</span><span class="sxs-lookup"><span data-stu-id="f348b-145">OWAView: String</span></span>

<span data-ttu-id="f348b-146">Получает строку, представляющую текущее представление Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="f348b-146">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="f348b-147">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="f348b-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="f348b-148">Если ведущее приложение не является Outlook в Интернете, то при доступе к этому свойству будет получен результат `undefined`.</span><span class="sxs-lookup"><span data-stu-id="f348b-148">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="f348b-149">В Outlook в Интернете есть три представления, которые соответствуют ширине экрана и окна, а также количество отображаемых столбцов:</span><span class="sxs-lookup"><span data-stu-id="f348b-149">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="f348b-150">`OneColumn`, который отображается, когда экран сужается.</span><span class="sxs-lookup"><span data-stu-id="f348b-150">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="f348b-151">В Outlook в Интернете этот макет с одним столбцом используется на всем экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="f348b-151">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="f348b-152">`TwoColumns`, который отображается, когда экран расширяется.</span><span class="sxs-lookup"><span data-stu-id="f348b-152">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="f348b-153">Outlook в Интернете использует это представление на большинстве планшетов.</span><span class="sxs-lookup"><span data-stu-id="f348b-153">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="f348b-154">`ThreeColumns` используется для полноразмерных экранов.</span><span class="sxs-lookup"><span data-stu-id="f348b-154">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="f348b-155">Например, в Outlook в Интернете это представление используется в полноэкранном окне на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="f348b-155">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="f348b-156">Тип</span><span class="sxs-lookup"><span data-stu-id="f348b-156">Type</span></span>

*   <span data-ttu-id="f348b-157">String</span><span class="sxs-lookup"><span data-stu-id="f348b-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f348b-158">Требования</span><span class="sxs-lookup"><span data-stu-id="f348b-158">Requirements</span></span>

|<span data-ttu-id="f348b-159">Требование</span><span class="sxs-lookup"><span data-stu-id="f348b-159">Requirement</span></span>| <span data-ttu-id="f348b-160">Значение</span><span class="sxs-lookup"><span data-stu-id="f348b-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="f348b-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f348b-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f348b-162">1.0</span><span class="sxs-lookup"><span data-stu-id="f348b-162">1.0</span></span>|
|[<span data-ttu-id="f348b-163">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f348b-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f348b-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f348b-164">ReadItem</span></span>|
|[<span data-ttu-id="f348b-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f348b-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f348b-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f348b-166">Compose or Read</span></span>|
