---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.4
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 7facc0ea555dca7d6784a09f798c3d8fa25f2731
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067848"
---
# <a name="userprofile"></a><span data-ttu-id="d4561-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="d4561-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="d4561-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="d4561-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4561-104">Требования</span><span class="sxs-lookup"><span data-stu-id="d4561-104">Requirements</span></span>

|<span data-ttu-id="d4561-105">Требование</span><span class="sxs-lookup"><span data-stu-id="d4561-105">Requirement</span></span>| <span data-ttu-id="d4561-106">Значение</span><span class="sxs-lookup"><span data-stu-id="d4561-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4561-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d4561-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4561-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d4561-108">1.0</span></span>|
|[<span data-ttu-id="d4561-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d4561-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4561-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4561-110">ReadItem</span></span>|
|[<span data-ttu-id="d4561-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d4561-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4561-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d4561-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="d4561-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="d4561-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="d4561-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="d4561-114">displayName :String</span></span>

<span data-ttu-id="d4561-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="d4561-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="d4561-116">Тип</span><span class="sxs-lookup"><span data-stu-id="d4561-116">Type</span></span>

*   <span data-ttu-id="d4561-117">String</span><span class="sxs-lookup"><span data-stu-id="d4561-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4561-118">Требования</span><span class="sxs-lookup"><span data-stu-id="d4561-118">Requirements</span></span>

|<span data-ttu-id="d4561-119">Требование</span><span class="sxs-lookup"><span data-stu-id="d4561-119">Requirement</span></span>| <span data-ttu-id="d4561-120">Значение</span><span class="sxs-lookup"><span data-stu-id="d4561-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4561-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d4561-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4561-122">1.0</span><span class="sxs-lookup"><span data-stu-id="d4561-122">1.0</span></span>|
|[<span data-ttu-id="d4561-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d4561-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4561-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4561-124">ReadItem</span></span>|
|[<span data-ttu-id="d4561-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d4561-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4561-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d4561-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4561-127">Пример</span><span class="sxs-lookup"><span data-stu-id="d4561-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="d4561-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="d4561-128">emailAddress :String</span></span>

<span data-ttu-id="d4561-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="d4561-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="d4561-130">Тип</span><span class="sxs-lookup"><span data-stu-id="d4561-130">Type</span></span>

*   <span data-ttu-id="d4561-131">String</span><span class="sxs-lookup"><span data-stu-id="d4561-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4561-132">Требования</span><span class="sxs-lookup"><span data-stu-id="d4561-132">Requirements</span></span>

|<span data-ttu-id="d4561-133">Требование</span><span class="sxs-lookup"><span data-stu-id="d4561-133">Requirement</span></span>| <span data-ttu-id="d4561-134">Значение</span><span class="sxs-lookup"><span data-stu-id="d4561-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4561-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d4561-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4561-136">1.0</span><span class="sxs-lookup"><span data-stu-id="d4561-136">1.0</span></span>|
|[<span data-ttu-id="d4561-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d4561-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4561-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4561-138">ReadItem</span></span>|
|[<span data-ttu-id="d4561-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d4561-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4561-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d4561-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4561-141">Пример</span><span class="sxs-lookup"><span data-stu-id="d4561-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="d4561-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="d4561-142">timeZone :String</span></span>

<span data-ttu-id="d4561-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="d4561-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="d4561-144">Тип</span><span class="sxs-lookup"><span data-stu-id="d4561-144">Type</span></span>

*   <span data-ttu-id="d4561-145">String</span><span class="sxs-lookup"><span data-stu-id="d4561-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4561-146">Требования</span><span class="sxs-lookup"><span data-stu-id="d4561-146">Requirements</span></span>

|<span data-ttu-id="d4561-147">Требование</span><span class="sxs-lookup"><span data-stu-id="d4561-147">Requirement</span></span>| <span data-ttu-id="d4561-148">Значение</span><span class="sxs-lookup"><span data-stu-id="d4561-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4561-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d4561-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4561-150">1.0</span><span class="sxs-lookup"><span data-stu-id="d4561-150">1.0</span></span>|
|[<span data-ttu-id="d4561-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d4561-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4561-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4561-152">ReadItem</span></span>|
|[<span data-ttu-id="d4561-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d4561-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4561-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d4561-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d4561-155">Пример</span><span class="sxs-lookup"><span data-stu-id="d4561-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
