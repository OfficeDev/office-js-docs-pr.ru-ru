---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2798b07b3353e9d89f757a22e6bed19dbd94a1c5
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450319"
---
# <a name="userprofile"></a><span data-ttu-id="9f2be-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="9f2be-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="9f2be-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="9f2be-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f2be-104">Требования</span><span class="sxs-lookup"><span data-stu-id="9f2be-104">Requirements</span></span>

|<span data-ttu-id="9f2be-105">Требование</span><span class="sxs-lookup"><span data-stu-id="9f2be-105">Requirement</span></span>| <span data-ttu-id="9f2be-106">Значение</span><span class="sxs-lookup"><span data-stu-id="9f2be-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f2be-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9f2be-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f2be-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9f2be-108">1.0</span></span>|
|[<span data-ttu-id="9f2be-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9f2be-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f2be-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f2be-110">ReadItem</span></span>|
|[<span data-ttu-id="9f2be-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9f2be-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f2be-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9f2be-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="9f2be-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="9f2be-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="9f2be-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="9f2be-114">displayName :String</span></span>

<span data-ttu-id="9f2be-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="9f2be-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="9f2be-116">Тип</span><span class="sxs-lookup"><span data-stu-id="9f2be-116">Type</span></span>

*   <span data-ttu-id="9f2be-117">String</span><span class="sxs-lookup"><span data-stu-id="9f2be-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f2be-118">Требования</span><span class="sxs-lookup"><span data-stu-id="9f2be-118">Requirements</span></span>

|<span data-ttu-id="9f2be-119">Требование</span><span class="sxs-lookup"><span data-stu-id="9f2be-119">Requirement</span></span>| <span data-ttu-id="9f2be-120">Значение</span><span class="sxs-lookup"><span data-stu-id="9f2be-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f2be-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9f2be-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f2be-122">1.0</span><span class="sxs-lookup"><span data-stu-id="9f2be-122">1.0</span></span>|
|[<span data-ttu-id="9f2be-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9f2be-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f2be-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f2be-124">ReadItem</span></span>|
|[<span data-ttu-id="9f2be-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9f2be-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f2be-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9f2be-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f2be-127">Пример</span><span class="sxs-lookup"><span data-stu-id="9f2be-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="9f2be-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="9f2be-128">emailAddress :String</span></span>

<span data-ttu-id="9f2be-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="9f2be-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="9f2be-130">Тип</span><span class="sxs-lookup"><span data-stu-id="9f2be-130">Type</span></span>

*   <span data-ttu-id="9f2be-131">String</span><span class="sxs-lookup"><span data-stu-id="9f2be-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f2be-132">Требования</span><span class="sxs-lookup"><span data-stu-id="9f2be-132">Requirements</span></span>

|<span data-ttu-id="9f2be-133">Требование</span><span class="sxs-lookup"><span data-stu-id="9f2be-133">Requirement</span></span>| <span data-ttu-id="9f2be-134">Значение</span><span class="sxs-lookup"><span data-stu-id="9f2be-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f2be-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9f2be-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f2be-136">1.0</span><span class="sxs-lookup"><span data-stu-id="9f2be-136">1.0</span></span>|
|[<span data-ttu-id="9f2be-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9f2be-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f2be-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f2be-138">ReadItem</span></span>|
|[<span data-ttu-id="9f2be-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9f2be-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f2be-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9f2be-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f2be-141">Пример</span><span class="sxs-lookup"><span data-stu-id="9f2be-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="9f2be-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="9f2be-142">timeZone :String</span></span>

<span data-ttu-id="9f2be-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="9f2be-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="9f2be-144">Тип</span><span class="sxs-lookup"><span data-stu-id="9f2be-144">Type</span></span>

*   <span data-ttu-id="9f2be-145">String</span><span class="sxs-lookup"><span data-stu-id="9f2be-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f2be-146">Требования</span><span class="sxs-lookup"><span data-stu-id="9f2be-146">Requirements</span></span>

|<span data-ttu-id="9f2be-147">Требование</span><span class="sxs-lookup"><span data-stu-id="9f2be-147">Requirement</span></span>| <span data-ttu-id="9f2be-148">Значение</span><span class="sxs-lookup"><span data-stu-id="9f2be-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f2be-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9f2be-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f2be-150">1.0</span><span class="sxs-lookup"><span data-stu-id="9f2be-150">1.0</span></span>|
|[<span data-ttu-id="9f2be-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9f2be-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f2be-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f2be-152">ReadItem</span></span>|
|[<span data-ttu-id="9f2be-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9f2be-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f2be-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9f2be-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9f2be-155">Пример</span><span class="sxs-lookup"><span data-stu-id="9f2be-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
