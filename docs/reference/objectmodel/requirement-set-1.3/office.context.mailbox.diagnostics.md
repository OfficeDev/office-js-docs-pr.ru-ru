---
title: Office.context.mailbox.diagnostics — набор обязательных элементов 1.3
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: a0edb17bebed0965b6fc642f386b94aaca2e4ad3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433133"
---
# <a name="diagnostics"></a><span data-ttu-id="155a6-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="155a6-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="155a6-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="155a6-103">Office.context.mailbox.diagnostics</span></span>

<span data-ttu-id="155a6-104">Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="155a6-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="155a6-105">Требования</span><span class="sxs-lookup"><span data-stu-id="155a6-105">Requirements</span></span>

|<span data-ttu-id="155a6-106">Требование</span><span class="sxs-lookup"><span data-stu-id="155a6-106">Requirement</span></span>| <span data-ttu-id="155a6-107">Значение</span><span class="sxs-lookup"><span data-stu-id="155a6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="155a6-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="155a6-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="155a6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="155a6-109">1.0</span></span>|
|[<span data-ttu-id="155a6-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="155a6-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="155a6-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="155a6-111">ReadItem</span></span>|
|[<span data-ttu-id="155a6-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="155a6-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="155a6-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="155a6-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="155a6-114">Элементы</span><span class="sxs-lookup"><span data-stu-id="155a6-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="155a6-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="155a6-115">hostName :String</span></span>

<span data-ttu-id="155a6-116">Получает строку, представляющую имя ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="155a6-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="155a6-117">Строка, которая может иметь одно из следующих значений: `Outlook`, `OutlookIOS` или `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="155a6-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, , or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="155a6-118">Тип:</span><span class="sxs-lookup"><span data-stu-id="155a6-118">Type:</span></span>

*   <span data-ttu-id="155a6-119">String</span><span class="sxs-lookup"><span data-stu-id="155a6-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="155a6-120">Требования</span><span class="sxs-lookup"><span data-stu-id="155a6-120">Requirements</span></span>

|<span data-ttu-id="155a6-121">Требование</span><span class="sxs-lookup"><span data-stu-id="155a6-121">Requirement</span></span>| <span data-ttu-id="155a6-122">Значение</span><span class="sxs-lookup"><span data-stu-id="155a6-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="155a6-123">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="155a6-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="155a6-124">1.0</span><span class="sxs-lookup"><span data-stu-id="155a6-124">1.0</span></span>|
|[<span data-ttu-id="155a6-125">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="155a6-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="155a6-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="155a6-126">ReadItem</span></span>|
|[<span data-ttu-id="155a6-127">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="155a6-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="155a6-128">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="155a6-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="155a6-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="155a6-129">hostVersion :String</span></span>

<span data-ttu-id="155a6-130">Получает строку, которая представляет версию ведущего приложения или Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="155a6-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="155a6-p101">Если почтовая надстройка запущена в классическом клиенте Outlook или Outlook для iOS, свойство `hostVersion` возвращает версию ведущего приложения, Outlook. В Outlook Web App это свойство возвращает версию Exchange Server. Пример — строка `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="155a6-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="155a6-134">Тип:</span><span class="sxs-lookup"><span data-stu-id="155a6-134">Type:</span></span>

*   <span data-ttu-id="155a6-135">String</span><span class="sxs-lookup"><span data-stu-id="155a6-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="155a6-136">Требования</span><span class="sxs-lookup"><span data-stu-id="155a6-136">Requirements</span></span>

|<span data-ttu-id="155a6-137">Требование</span><span class="sxs-lookup"><span data-stu-id="155a6-137">Requirement</span></span>| <span data-ttu-id="155a6-138">Значение</span><span class="sxs-lookup"><span data-stu-id="155a6-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="155a6-139">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="155a6-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="155a6-140">1.0</span><span class="sxs-lookup"><span data-stu-id="155a6-140">1.0</span></span>|
|[<span data-ttu-id="155a6-141">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="155a6-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="155a6-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="155a6-142">ReadItem</span></span>|
|[<span data-ttu-id="155a6-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="155a6-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="155a6-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="155a6-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="155a6-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="155a6-145">OWAView :String</span></span>

<span data-ttu-id="155a6-146">Получает строку, отображающую текущее представление Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="155a6-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="155a6-147">Возвращаемая строка может иметь одно из следующих значений: `OneColumn`, `TwoColumns` или `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="155a6-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="155a6-148">Если Outlook Web App — не ведущее приложение, при получении доступа к этому свойству будет выдаваться значение `undefined`.</span><span class="sxs-lookup"><span data-stu-id="155a6-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="155a6-149">Outlook Web App включает три представления, которые соответствуют ширине экрана и окна, а также числу отображаемых столбцов.</span><span class="sxs-lookup"><span data-stu-id="155a6-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="155a6-p102">`OneColumn` используется в случае узкого экрана: Outlook Web App использует этот макет размером в один столбец на экране смартфона.</span><span class="sxs-lookup"><span data-stu-id="155a6-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="155a6-p103">`TwoColumns` используется при более широком экране: Outlook Web App использует это представление на большинстве планшетных ПК.</span><span class="sxs-lookup"><span data-stu-id="155a6-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="155a6-p104">`ThreeColumns` используется для полноразмерных экранов. Например, Outlook Web App использует это представление в полноэкранном режиме на настольных компьютерах.</span><span class="sxs-lookup"><span data-stu-id="155a6-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="155a6-156">Тип:</span><span class="sxs-lookup"><span data-stu-id="155a6-156">Type:</span></span>

*   <span data-ttu-id="155a6-157">String</span><span class="sxs-lookup"><span data-stu-id="155a6-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="155a6-158">Требования</span><span class="sxs-lookup"><span data-stu-id="155a6-158">Requirements</span></span>

|<span data-ttu-id="155a6-159">Требование</span><span class="sxs-lookup"><span data-stu-id="155a6-159">Requirement</span></span>| <span data-ttu-id="155a6-160">Значение</span><span class="sxs-lookup"><span data-stu-id="155a6-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="155a6-161">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="155a6-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="155a6-162">1.0</span><span class="sxs-lookup"><span data-stu-id="155a6-162">1.0</span></span>|
|[<span data-ttu-id="155a6-163">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="155a6-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="155a6-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="155a6-164">ReadItem</span></span>|
|[<span data-ttu-id="155a6-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="155a6-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="155a6-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="155a6-166">Compose or read</span></span>|