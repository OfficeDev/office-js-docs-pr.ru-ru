---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,2
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 496a59f4ef02f03cda95fde0bf14634b1db13f77
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870585"
---
# <a name="userprofile"></a><span data-ttu-id="dd70f-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="dd70f-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="dd70f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="dd70f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd70f-104">Требования</span><span class="sxs-lookup"><span data-stu-id="dd70f-104">Requirements</span></span>

|<span data-ttu-id="dd70f-105">Требование</span><span class="sxs-lookup"><span data-stu-id="dd70f-105">Requirement</span></span>| <span data-ttu-id="dd70f-106">Значение</span><span class="sxs-lookup"><span data-stu-id="dd70f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70f-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd70f-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70f-108">1.0</span></span>|
|[<span data-ttu-id="dd70f-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd70f-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70f-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70f-110">ReadItem</span></span>|
|[<span data-ttu-id="dd70f-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd70f-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70f-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd70f-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="dd70f-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="dd70f-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="dd70f-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="dd70f-114">displayName :String</span></span>

<span data-ttu-id="dd70f-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="dd70f-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="dd70f-116">Тип</span><span class="sxs-lookup"><span data-stu-id="dd70f-116">Type</span></span>

*   <span data-ttu-id="dd70f-117">String</span><span class="sxs-lookup"><span data-stu-id="dd70f-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd70f-118">Требования</span><span class="sxs-lookup"><span data-stu-id="dd70f-118">Requirements</span></span>

|<span data-ttu-id="dd70f-119">Требование</span><span class="sxs-lookup"><span data-stu-id="dd70f-119">Requirement</span></span>| <span data-ttu-id="dd70f-120">Значение</span><span class="sxs-lookup"><span data-stu-id="dd70f-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70f-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd70f-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70f-122">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70f-122">1.0</span></span>|
|[<span data-ttu-id="dd70f-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd70f-123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70f-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70f-124">ReadItem</span></span>|
|[<span data-ttu-id="dd70f-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd70f-125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70f-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd70f-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd70f-127">Пример</span><span class="sxs-lookup"><span data-stu-id="dd70f-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="dd70f-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="dd70f-128">emailAddress :String</span></span>

<span data-ttu-id="dd70f-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="dd70f-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="dd70f-130">Тип</span><span class="sxs-lookup"><span data-stu-id="dd70f-130">Type</span></span>

*   <span data-ttu-id="dd70f-131">String</span><span class="sxs-lookup"><span data-stu-id="dd70f-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd70f-132">Требования</span><span class="sxs-lookup"><span data-stu-id="dd70f-132">Requirements</span></span>

|<span data-ttu-id="dd70f-133">Требование</span><span class="sxs-lookup"><span data-stu-id="dd70f-133">Requirement</span></span>| <span data-ttu-id="dd70f-134">Значение</span><span class="sxs-lookup"><span data-stu-id="dd70f-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70f-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd70f-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70f-136">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70f-136">1.0</span></span>|
|[<span data-ttu-id="dd70f-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd70f-137">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70f-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70f-138">ReadItem</span></span>|
|[<span data-ttu-id="dd70f-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd70f-139">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70f-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd70f-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd70f-141">Пример</span><span class="sxs-lookup"><span data-stu-id="dd70f-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="dd70f-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="dd70f-142">timeZone :String</span></span>

<span data-ttu-id="dd70f-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="dd70f-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="dd70f-144">Тип</span><span class="sxs-lookup"><span data-stu-id="dd70f-144">Type</span></span>

*   <span data-ttu-id="dd70f-145">String</span><span class="sxs-lookup"><span data-stu-id="dd70f-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd70f-146">Требования</span><span class="sxs-lookup"><span data-stu-id="dd70f-146">Requirements</span></span>

|<span data-ttu-id="dd70f-147">Требование</span><span class="sxs-lookup"><span data-stu-id="dd70f-147">Requirement</span></span>| <span data-ttu-id="dd70f-148">Значение</span><span class="sxs-lookup"><span data-stu-id="dd70f-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70f-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="dd70f-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70f-150">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70f-150">1.0</span></span>|
|[<span data-ttu-id="dd70f-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="dd70f-151">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70f-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70f-152">ReadItem</span></span>|
|[<span data-ttu-id="dd70f-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="dd70f-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70f-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="dd70f-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd70f-155">Пример</span><span class="sxs-lookup"><span data-stu-id="dd70f-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
