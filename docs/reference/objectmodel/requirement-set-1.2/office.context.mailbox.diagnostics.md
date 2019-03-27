---
title: Office. Context. Mailbox. Diagnostics — набор обязательных элементов 1,2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2b57a26233ce76ed35ce9428cc12a4ba93b5ace8
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871502"
---
# <a name="diagnostics"></a><span data-ttu-id="911d2-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="911d2-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="911d2-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="911d2-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="911d2-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="911d2-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="911d2-105">Требования</span><span class="sxs-lookup"><span data-stu-id="911d2-105">Requirements</span></span>

|<span data-ttu-id="911d2-106">Требование</span><span class="sxs-lookup"><span data-stu-id="911d2-106">Requirement</span></span>| <span data-ttu-id="911d2-107">Значение</span><span class="sxs-lookup"><span data-stu-id="911d2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="911d2-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="911d2-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="911d2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="911d2-109">1.0</span></span>|
|[<span data-ttu-id="911d2-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="911d2-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="911d2-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="911d2-111">ReadItem</span></span>|
|[<span data-ttu-id="911d2-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="911d2-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="911d2-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="911d2-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="911d2-114">Элементы</span><span class="sxs-lookup"><span data-stu-id="911d2-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="911d2-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="911d2-115">hostName :String</span></span>

<span data-ttu-id="911d2-116">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="911d2-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="911d2-117">Строка, которая может иметь одно из следующих значений: `Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="911d2-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="911d2-118">Тип</span><span class="sxs-lookup"><span data-stu-id="911d2-118">Type</span></span>

*   <span data-ttu-id="911d2-119">String</span><span class="sxs-lookup"><span data-stu-id="911d2-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="911d2-120">Требования</span><span class="sxs-lookup"><span data-stu-id="911d2-120">Requirements</span></span>

|<span data-ttu-id="911d2-121">Требование</span><span class="sxs-lookup"><span data-stu-id="911d2-121">Requirement</span></span>| <span data-ttu-id="911d2-122">Значение</span><span class="sxs-lookup"><span data-stu-id="911d2-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="911d2-123">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="911d2-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="911d2-124">1.0</span><span class="sxs-lookup"><span data-stu-id="911d2-124">1.0</span></span>|
|[<span data-ttu-id="911d2-125">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="911d2-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="911d2-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="911d2-126">ReadItem</span></span>|
|[<span data-ttu-id="911d2-127">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="911d2-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="911d2-128">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="911d2-128">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="911d2-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="911d2-129">hostVersion :String</span></span>

<span data-ttu-id="911d2-130">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="911d2-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="911d2-p101">Если почтовая надстройка запущена в классическом клиенте Outlook или Outlook для iOS, свойство `hostVersion` возвращает версию ведущего приложения, Outlook. В Outlook Web App это свойство возвращает версию Exchange Server. Пример — строка `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="911d2-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="911d2-134">Тип</span><span class="sxs-lookup"><span data-stu-id="911d2-134">Type</span></span>

*   <span data-ttu-id="911d2-135">String</span><span class="sxs-lookup"><span data-stu-id="911d2-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="911d2-136">Требования</span><span class="sxs-lookup"><span data-stu-id="911d2-136">Requirements</span></span>

|<span data-ttu-id="911d2-137">Требование</span><span class="sxs-lookup"><span data-stu-id="911d2-137">Requirement</span></span>| <span data-ttu-id="911d2-138">Значение</span><span class="sxs-lookup"><span data-stu-id="911d2-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="911d2-139">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="911d2-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="911d2-140">1.0</span><span class="sxs-lookup"><span data-stu-id="911d2-140">1.0</span></span>|
|[<span data-ttu-id="911d2-141">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="911d2-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="911d2-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="911d2-142">ReadItem</span></span>|
|[<span data-ttu-id="911d2-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="911d2-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="911d2-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="911d2-144">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="911d2-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="911d2-145">OWAView :String</span></span>

<span data-ttu-id="911d2-146">Получает строку, отображающую текущее представление Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="911d2-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="911d2-147">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="911d2-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="911d2-148">Если Outlook Web App — не ведущее приложение, при получении доступа к этому свойству будет выдаваться значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="911d2-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="911d2-149">Outlook Web App включает три представления, которые соответствуют ширине экрана и окна, а также числу отображаемых столбцов.</span><span class="sxs-lookup"><span data-stu-id="911d2-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="911d2-p102">`OneColumn` используется в случае узкого экрана: Outlook Web App использует этот макет размером в один столбец на экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="911d2-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="911d2-p103">`TwoColumns` используется при более широком экране: Outlook Web App использует это представление на большинстве планшетных ПК.</span><span class="sxs-lookup"><span data-stu-id="911d2-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="911d2-p104">`ThreeColumns` используется для полноразмерных экранов. Например, Outlook Web App использует это представление в полноэкранном режиме на настольных компьютерах.</span><span class="sxs-lookup"><span data-stu-id="911d2-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="911d2-156">Тип</span><span class="sxs-lookup"><span data-stu-id="911d2-156">Type</span></span>

*   <span data-ttu-id="911d2-157">String</span><span class="sxs-lookup"><span data-stu-id="911d2-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="911d2-158">Требования</span><span class="sxs-lookup"><span data-stu-id="911d2-158">Requirements</span></span>

|<span data-ttu-id="911d2-159">Требование</span><span class="sxs-lookup"><span data-stu-id="911d2-159">Requirement</span></span>| <span data-ttu-id="911d2-160">Значение</span><span class="sxs-lookup"><span data-stu-id="911d2-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="911d2-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="911d2-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="911d2-162">1.0</span><span class="sxs-lookup"><span data-stu-id="911d2-162">1.0</span></span>|
|[<span data-ttu-id="911d2-163">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="911d2-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="911d2-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="911d2-164">ReadItem</span></span>|
|[<span data-ttu-id="911d2-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="911d2-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="911d2-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="911d2-166">Compose or Read</span></span>|
