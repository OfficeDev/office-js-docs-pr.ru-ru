---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,3
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 03cdc13845bff0fbd3855f29f43298cd770e5ad9
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451845"
---
# <a name="userprofile"></a><span data-ttu-id="48c1e-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="48c1e-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="48c1e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="48c1e-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="48c1e-104">Требования</span><span class="sxs-lookup"><span data-stu-id="48c1e-104">Requirements</span></span>

|<span data-ttu-id="48c1e-105">Требование</span><span class="sxs-lookup"><span data-stu-id="48c1e-105">Requirement</span></span>| <span data-ttu-id="48c1e-106">Значение</span><span class="sxs-lookup"><span data-stu-id="48c1e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="48c1e-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48c1e-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48c1e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="48c1e-108">1.0</span></span>|
|[<span data-ttu-id="48c1e-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48c1e-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48c1e-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48c1e-110">ReadItem</span></span>|
|[<span data-ttu-id="48c1e-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48c1e-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="48c1e-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48c1e-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="48c1e-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="48c1e-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="48c1e-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="48c1e-114">displayName :String</span></span>

<span data-ttu-id="48c1e-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="48c1e-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="48c1e-116">Тип</span><span class="sxs-lookup"><span data-stu-id="48c1e-116">Type</span></span>

*   <span data-ttu-id="48c1e-117">String</span><span class="sxs-lookup"><span data-stu-id="48c1e-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48c1e-118">Требования</span><span class="sxs-lookup"><span data-stu-id="48c1e-118">Requirements</span></span>

|<span data-ttu-id="48c1e-119">Требование</span><span class="sxs-lookup"><span data-stu-id="48c1e-119">Requirement</span></span>| <span data-ttu-id="48c1e-120">Значение</span><span class="sxs-lookup"><span data-stu-id="48c1e-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="48c1e-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48c1e-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48c1e-122">1.0</span><span class="sxs-lookup"><span data-stu-id="48c1e-122">1.0</span></span>|
|[<span data-ttu-id="48c1e-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48c1e-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48c1e-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48c1e-124">ReadItem</span></span>|
|[<span data-ttu-id="48c1e-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48c1e-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="48c1e-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48c1e-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48c1e-127">Пример</span><span class="sxs-lookup"><span data-stu-id="48c1e-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="48c1e-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="48c1e-128">emailAddress :String</span></span>

<span data-ttu-id="48c1e-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="48c1e-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="48c1e-130">Тип</span><span class="sxs-lookup"><span data-stu-id="48c1e-130">Type</span></span>

*   <span data-ttu-id="48c1e-131">String</span><span class="sxs-lookup"><span data-stu-id="48c1e-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48c1e-132">Требования</span><span class="sxs-lookup"><span data-stu-id="48c1e-132">Requirements</span></span>

|<span data-ttu-id="48c1e-133">Требование</span><span class="sxs-lookup"><span data-stu-id="48c1e-133">Requirement</span></span>| <span data-ttu-id="48c1e-134">Значение</span><span class="sxs-lookup"><span data-stu-id="48c1e-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="48c1e-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48c1e-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48c1e-136">1.0</span><span class="sxs-lookup"><span data-stu-id="48c1e-136">1.0</span></span>|
|[<span data-ttu-id="48c1e-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48c1e-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48c1e-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48c1e-138">ReadItem</span></span>|
|[<span data-ttu-id="48c1e-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48c1e-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="48c1e-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48c1e-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48c1e-141">Пример</span><span class="sxs-lookup"><span data-stu-id="48c1e-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="48c1e-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="48c1e-142">timeZone :String</span></span>

<span data-ttu-id="48c1e-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="48c1e-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="48c1e-144">Тип</span><span class="sxs-lookup"><span data-stu-id="48c1e-144">Type</span></span>

*   <span data-ttu-id="48c1e-145">String</span><span class="sxs-lookup"><span data-stu-id="48c1e-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48c1e-146">Требования</span><span class="sxs-lookup"><span data-stu-id="48c1e-146">Requirements</span></span>

|<span data-ttu-id="48c1e-147">Требование</span><span class="sxs-lookup"><span data-stu-id="48c1e-147">Requirement</span></span>| <span data-ttu-id="48c1e-148">Значение</span><span class="sxs-lookup"><span data-stu-id="48c1e-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="48c1e-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="48c1e-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48c1e-150">1.0</span><span class="sxs-lookup"><span data-stu-id="48c1e-150">1.0</span></span>|
|[<span data-ttu-id="48c1e-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="48c1e-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48c1e-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48c1e-152">ReadItem</span></span>|
|[<span data-ttu-id="48c1e-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="48c1e-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="48c1e-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="48c1e-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48c1e-155">Пример</span><span class="sxs-lookup"><span data-stu-id="48c1e-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
