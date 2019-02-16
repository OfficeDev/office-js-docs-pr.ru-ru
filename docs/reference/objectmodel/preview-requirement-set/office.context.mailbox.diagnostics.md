---
title: Office.context.mailbox.diagnostics — предварительная версия набора обязательных элементов
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: edeb0ce8a19dae4ec12705f81c216d4d7857ae90
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068170"
---
# <a name="diagnostics"></a><span data-ttu-id="a8220-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="a8220-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="a8220-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="a8220-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="a8220-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="a8220-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a8220-105">Требования</span><span class="sxs-lookup"><span data-stu-id="a8220-105">Requirements</span></span>

|<span data-ttu-id="a8220-106">Требование</span><span class="sxs-lookup"><span data-stu-id="a8220-106">Requirement</span></span>| <span data-ttu-id="a8220-107">Значение</span><span class="sxs-lookup"><span data-stu-id="a8220-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8220-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a8220-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8220-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a8220-109">1.0</span></span>|
|[<span data-ttu-id="a8220-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a8220-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a8220-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a8220-111">ReadItem</span></span>|
|[<span data-ttu-id="a8220-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a8220-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8220-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a8220-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a8220-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="a8220-114">Members and methods</span></span>

| <span data-ttu-id="a8220-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="a8220-115">Member</span></span> | <span data-ttu-id="a8220-116">Тип</span><span class="sxs-lookup"><span data-stu-id="a8220-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a8220-117">hostName</span><span class="sxs-lookup"><span data-stu-id="a8220-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="a8220-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="a8220-118">Member</span></span> |
| [<span data-ttu-id="a8220-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="a8220-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="a8220-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="a8220-120">Member</span></span> |
| [<span data-ttu-id="a8220-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="a8220-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="a8220-122">Член</span><span class="sxs-lookup"><span data-stu-id="a8220-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="a8220-123">Элементы</span><span class="sxs-lookup"><span data-stu-id="a8220-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="a8220-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="a8220-124">hostName :String</span></span>

<span data-ttu-id="a8220-125">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="a8220-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="a8220-126">Строка, которая может иметь одно из следующих значений: `Outlook`, `Mac Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="a8220-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="a8220-127">Тип</span><span class="sxs-lookup"><span data-stu-id="a8220-127">Type</span></span>

*   <span data-ttu-id="a8220-128">String</span><span class="sxs-lookup"><span data-stu-id="a8220-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a8220-129">Требования</span><span class="sxs-lookup"><span data-stu-id="a8220-129">Requirements</span></span>

|<span data-ttu-id="a8220-130">Требование</span><span class="sxs-lookup"><span data-stu-id="a8220-130">Requirement</span></span>| <span data-ttu-id="a8220-131">Значение</span><span class="sxs-lookup"><span data-stu-id="a8220-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8220-132">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a8220-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8220-133">1.0</span><span class="sxs-lookup"><span data-stu-id="a8220-133">1.0</span></span>|
|[<span data-ttu-id="a8220-134">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a8220-134">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a8220-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a8220-135">ReadItem</span></span>|
|[<span data-ttu-id="a8220-136">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a8220-136">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8220-137">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a8220-137">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="a8220-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="a8220-138">hostVersion :String</span></span>

<span data-ttu-id="a8220-139">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="a8220-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="a8220-p101">Если почтовая надстройка запущена в классическом клиенте Outlook или Outlook для iOS, свойство `hostVersion` возвращает версию ведущего приложения, Outlook. В Outlook Web App это свойство возвращает версию Exchange Server. Пример — строка `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="a8220-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="a8220-143">Тип</span><span class="sxs-lookup"><span data-stu-id="a8220-143">Type</span></span>

*   <span data-ttu-id="a8220-144">String</span><span class="sxs-lookup"><span data-stu-id="a8220-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a8220-145">Требования</span><span class="sxs-lookup"><span data-stu-id="a8220-145">Requirements</span></span>

|<span data-ttu-id="a8220-146">Требование</span><span class="sxs-lookup"><span data-stu-id="a8220-146">Requirement</span></span>| <span data-ttu-id="a8220-147">Значение</span><span class="sxs-lookup"><span data-stu-id="a8220-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8220-148">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a8220-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8220-149">1.0</span><span class="sxs-lookup"><span data-stu-id="a8220-149">1.0</span></span>|
|[<span data-ttu-id="a8220-150">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a8220-150">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a8220-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a8220-151">ReadItem</span></span>|
|[<span data-ttu-id="a8220-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a8220-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8220-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a8220-153">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="a8220-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="a8220-154">OWAView :String</span></span>

<span data-ttu-id="a8220-155">Получает строку, отображающую текущее представление Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="a8220-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="a8220-156">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="a8220-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="a8220-157">Если Outlook Web App — не ведущее приложение, при получении доступа к этому свойству будет выдаваться значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="a8220-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="a8220-158">Outlook Web App включает три представления, которые соответствуют ширине экрана и окна, а также числу отображаемых столбцов.</span><span class="sxs-lookup"><span data-stu-id="a8220-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="a8220-p102">`OneColumn` используется в случае узкого экрана: Outlook Web App использует этот макет размером в один столбец на экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="a8220-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="a8220-p103">`TwoColumns` используется при более широком экране: Outlook Web App использует это представление на большинстве планшетных ПК.</span><span class="sxs-lookup"><span data-stu-id="a8220-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="a8220-p104">`ThreeColumns` используется для полноразмерных экранов. Например, Outlook Web App использует это представление в полноэкранном режиме на настольных компьютерах.</span><span class="sxs-lookup"><span data-stu-id="a8220-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="a8220-165">Тип</span><span class="sxs-lookup"><span data-stu-id="a8220-165">Type</span></span>

*   <span data-ttu-id="a8220-166">String</span><span class="sxs-lookup"><span data-stu-id="a8220-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a8220-167">Требования</span><span class="sxs-lookup"><span data-stu-id="a8220-167">Requirements</span></span>

|<span data-ttu-id="a8220-168">Требование</span><span class="sxs-lookup"><span data-stu-id="a8220-168">Requirement</span></span>| <span data-ttu-id="a8220-169">Значение</span><span class="sxs-lookup"><span data-stu-id="a8220-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8220-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a8220-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8220-171">1.0</span><span class="sxs-lookup"><span data-stu-id="a8220-171">1.0</span></span>|
|[<span data-ttu-id="a8220-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a8220-172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a8220-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a8220-173">ReadItem</span></span>|
|[<span data-ttu-id="a8220-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a8220-174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8220-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a8220-175">Compose or Read</span></span>|
