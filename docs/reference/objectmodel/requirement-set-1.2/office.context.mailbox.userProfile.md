---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 496a59f4ef02f03cda95fde0bf14634b1db13f77
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450340"
---
# <a name="userprofile"></a><span data-ttu-id="53185-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="53185-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="53185-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="53185-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="53185-104">Требования</span><span class="sxs-lookup"><span data-stu-id="53185-104">Requirements</span></span>

|<span data-ttu-id="53185-105">Требование</span><span class="sxs-lookup"><span data-stu-id="53185-105">Requirement</span></span>| <span data-ttu-id="53185-106">Значение</span><span class="sxs-lookup"><span data-stu-id="53185-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="53185-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53185-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53185-108">1.0</span><span class="sxs-lookup"><span data-stu-id="53185-108">1.0</span></span>|
|[<span data-ttu-id="53185-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="53185-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53185-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53185-110">ReadItem</span></span>|
|[<span data-ttu-id="53185-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53185-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53185-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53185-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="53185-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="53185-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="53185-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="53185-114">displayName :String</span></span>

<span data-ttu-id="53185-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="53185-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="53185-116">Тип</span><span class="sxs-lookup"><span data-stu-id="53185-116">Type</span></span>

*   <span data-ttu-id="53185-117">String</span><span class="sxs-lookup"><span data-stu-id="53185-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="53185-118">Требования</span><span class="sxs-lookup"><span data-stu-id="53185-118">Requirements</span></span>

|<span data-ttu-id="53185-119">Требование</span><span class="sxs-lookup"><span data-stu-id="53185-119">Requirement</span></span>| <span data-ttu-id="53185-120">Значение</span><span class="sxs-lookup"><span data-stu-id="53185-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="53185-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53185-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53185-122">1.0</span><span class="sxs-lookup"><span data-stu-id="53185-122">1.0</span></span>|
|[<span data-ttu-id="53185-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="53185-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53185-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53185-124">ReadItem</span></span>|
|[<span data-ttu-id="53185-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53185-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53185-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53185-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="53185-127">Пример</span><span class="sxs-lookup"><span data-stu-id="53185-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="53185-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="53185-128">emailAddress :String</span></span>

<span data-ttu-id="53185-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="53185-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="53185-130">Тип</span><span class="sxs-lookup"><span data-stu-id="53185-130">Type</span></span>

*   <span data-ttu-id="53185-131">String</span><span class="sxs-lookup"><span data-stu-id="53185-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="53185-132">Требования</span><span class="sxs-lookup"><span data-stu-id="53185-132">Requirements</span></span>

|<span data-ttu-id="53185-133">Требование</span><span class="sxs-lookup"><span data-stu-id="53185-133">Requirement</span></span>| <span data-ttu-id="53185-134">Значение</span><span class="sxs-lookup"><span data-stu-id="53185-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="53185-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53185-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53185-136">1.0</span><span class="sxs-lookup"><span data-stu-id="53185-136">1.0</span></span>|
|[<span data-ttu-id="53185-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="53185-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53185-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53185-138">ReadItem</span></span>|
|[<span data-ttu-id="53185-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53185-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53185-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53185-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="53185-141">Пример</span><span class="sxs-lookup"><span data-stu-id="53185-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="53185-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="53185-142">timeZone :String</span></span>

<span data-ttu-id="53185-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="53185-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="53185-144">Тип</span><span class="sxs-lookup"><span data-stu-id="53185-144">Type</span></span>

*   <span data-ttu-id="53185-145">String</span><span class="sxs-lookup"><span data-stu-id="53185-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="53185-146">Требования</span><span class="sxs-lookup"><span data-stu-id="53185-146">Requirements</span></span>

|<span data-ttu-id="53185-147">Требование</span><span class="sxs-lookup"><span data-stu-id="53185-147">Requirement</span></span>| <span data-ttu-id="53185-148">Значение</span><span class="sxs-lookup"><span data-stu-id="53185-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="53185-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="53185-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="53185-150">1.0</span><span class="sxs-lookup"><span data-stu-id="53185-150">1.0</span></span>|
|[<span data-ttu-id="53185-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="53185-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="53185-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="53185-152">ReadItem</span></span>|
|[<span data-ttu-id="53185-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="53185-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53185-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="53185-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="53185-155">Пример</span><span class="sxs-lookup"><span data-stu-id="53185-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
