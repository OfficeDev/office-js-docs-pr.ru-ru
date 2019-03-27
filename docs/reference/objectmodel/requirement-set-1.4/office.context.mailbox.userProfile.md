---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 2798b07b3353e9d89f757a22e6bed19dbd94a1c5
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870046"
---
# <a name="userprofile"></a><span data-ttu-id="ce794-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ce794-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ce794-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ce794-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce794-104">Требования</span><span class="sxs-lookup"><span data-stu-id="ce794-104">Requirements</span></span>

|<span data-ttu-id="ce794-105">Требование</span><span class="sxs-lookup"><span data-stu-id="ce794-105">Requirement</span></span>| <span data-ttu-id="ce794-106">Значение</span><span class="sxs-lookup"><span data-stu-id="ce794-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce794-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce794-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce794-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ce794-108">1.0</span></span>|
|[<span data-ttu-id="ce794-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce794-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce794-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce794-110">ReadItem</span></span>|
|[<span data-ttu-id="ce794-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce794-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce794-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce794-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="ce794-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="ce794-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="ce794-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ce794-114">displayName :String</span></span>

<span data-ttu-id="ce794-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="ce794-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ce794-116">Тип</span><span class="sxs-lookup"><span data-stu-id="ce794-116">Type</span></span>

*   <span data-ttu-id="ce794-117">String</span><span class="sxs-lookup"><span data-stu-id="ce794-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce794-118">Требования</span><span class="sxs-lookup"><span data-stu-id="ce794-118">Requirements</span></span>

|<span data-ttu-id="ce794-119">Требование</span><span class="sxs-lookup"><span data-stu-id="ce794-119">Requirement</span></span>| <span data-ttu-id="ce794-120">Значение</span><span class="sxs-lookup"><span data-stu-id="ce794-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce794-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce794-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce794-122">1.0</span><span class="sxs-lookup"><span data-stu-id="ce794-122">1.0</span></span>|
|[<span data-ttu-id="ce794-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce794-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce794-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce794-124">ReadItem</span></span>|
|[<span data-ttu-id="ce794-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce794-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce794-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce794-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce794-127">Пример</span><span class="sxs-lookup"><span data-stu-id="ce794-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ce794-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ce794-128">emailAddress :String</span></span>

<span data-ttu-id="ce794-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="ce794-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ce794-130">Тип</span><span class="sxs-lookup"><span data-stu-id="ce794-130">Type</span></span>

*   <span data-ttu-id="ce794-131">String</span><span class="sxs-lookup"><span data-stu-id="ce794-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce794-132">Требования</span><span class="sxs-lookup"><span data-stu-id="ce794-132">Requirements</span></span>

|<span data-ttu-id="ce794-133">Требование</span><span class="sxs-lookup"><span data-stu-id="ce794-133">Requirement</span></span>| <span data-ttu-id="ce794-134">Значение</span><span class="sxs-lookup"><span data-stu-id="ce794-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce794-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce794-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce794-136">1.0</span><span class="sxs-lookup"><span data-stu-id="ce794-136">1.0</span></span>|
|[<span data-ttu-id="ce794-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce794-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce794-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce794-138">ReadItem</span></span>|
|[<span data-ttu-id="ce794-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce794-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce794-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce794-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce794-141">Пример</span><span class="sxs-lookup"><span data-stu-id="ce794-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ce794-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ce794-142">timeZone :String</span></span>

<span data-ttu-id="ce794-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ce794-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ce794-144">Тип</span><span class="sxs-lookup"><span data-stu-id="ce794-144">Type</span></span>

*   <span data-ttu-id="ce794-145">String</span><span class="sxs-lookup"><span data-stu-id="ce794-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce794-146">Требования</span><span class="sxs-lookup"><span data-stu-id="ce794-146">Requirements</span></span>

|<span data-ttu-id="ce794-147">Требование</span><span class="sxs-lookup"><span data-stu-id="ce794-147">Requirement</span></span>| <span data-ttu-id="ce794-148">Значение</span><span class="sxs-lookup"><span data-stu-id="ce794-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce794-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce794-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce794-150">1.0</span><span class="sxs-lookup"><span data-stu-id="ce794-150">1.0</span></span>|
|[<span data-ttu-id="ce794-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce794-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce794-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce794-152">ReadItem</span></span>|
|[<span data-ttu-id="ce794-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce794-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce794-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce794-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce794-155">Пример</span><span class="sxs-lookup"><span data-stu-id="ce794-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
