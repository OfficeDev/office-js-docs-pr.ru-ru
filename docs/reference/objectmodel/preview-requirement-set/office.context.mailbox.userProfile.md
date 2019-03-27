---
title: Office. Context. Mailbox. userProfile — Предварительная версия набора требований
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 204097497c958c26a6e67fc01d6dbd5142d8dced
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871873"
---
# <a name="userprofile"></a><span data-ttu-id="554b9-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="554b9-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="554b9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="554b9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="554b9-104">Требования</span><span class="sxs-lookup"><span data-stu-id="554b9-104">Requirements</span></span>

|<span data-ttu-id="554b9-105">Требование</span><span class="sxs-lookup"><span data-stu-id="554b9-105">Requirement</span></span>| <span data-ttu-id="554b9-106">Значение</span><span class="sxs-lookup"><span data-stu-id="554b9-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="554b9-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="554b9-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="554b9-108">1.0</span><span class="sxs-lookup"><span data-stu-id="554b9-108">1.0</span></span>|
|[<span data-ttu-id="554b9-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="554b9-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="554b9-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="554b9-110">ReadItem</span></span>|
|[<span data-ttu-id="554b9-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="554b9-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="554b9-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="554b9-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="554b9-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="554b9-113">Members and methods</span></span>

| <span data-ttu-id="554b9-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="554b9-114">Member</span></span> | <span data-ttu-id="554b9-115">Тип</span><span class="sxs-lookup"><span data-stu-id="554b9-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="554b9-116">accountType</span><span class="sxs-lookup"><span data-stu-id="554b9-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="554b9-117">Member</span><span class="sxs-lookup"><span data-stu-id="554b9-117">Member</span></span> |
| [<span data-ttu-id="554b9-118">displayName</span><span class="sxs-lookup"><span data-stu-id="554b9-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="554b9-119">Member</span><span class="sxs-lookup"><span data-stu-id="554b9-119">Member</span></span> |
| [<span data-ttu-id="554b9-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="554b9-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="554b9-121">Member</span><span class="sxs-lookup"><span data-stu-id="554b9-121">Member</span></span> |
| [<span data-ttu-id="554b9-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="554b9-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="554b9-123">Member</span><span class="sxs-lookup"><span data-stu-id="554b9-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="554b9-124">Members</span><span class="sxs-lookup"><span data-stu-id="554b9-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="554b9-125">accountType: строка</span><span class="sxs-lookup"><span data-stu-id="554b9-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="554b9-126">В настоящее время этот элемент поддерживается только в Outlook 2016 или более поздней версии для Mac (сборка 16.9.1212 или более поздняя).</span><span class="sxs-lookup"><span data-stu-id="554b9-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="554b9-127">Возвращает тип учетной записи пользователя, связанного с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="554b9-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="554b9-128">Возможные значения перечислены в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="554b9-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="554b9-129">Значение</span><span class="sxs-lookup"><span data-stu-id="554b9-129">Value</span></span> | <span data-ttu-id="554b9-130">Описание</span><span class="sxs-lookup"><span data-stu-id="554b9-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="554b9-131">Почтовый ящик находится на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="554b9-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="554b9-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="554b9-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="554b9-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="554b9-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="554b9-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="554b9-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="554b9-135">Тип</span><span class="sxs-lookup"><span data-stu-id="554b9-135">Type</span></span>

*   <span data-ttu-id="554b9-136">String</span><span class="sxs-lookup"><span data-stu-id="554b9-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="554b9-137">Требования</span><span class="sxs-lookup"><span data-stu-id="554b9-137">Requirements</span></span>

|<span data-ttu-id="554b9-138">Требование</span><span class="sxs-lookup"><span data-stu-id="554b9-138">Requirement</span></span>| <span data-ttu-id="554b9-139">Значение</span><span class="sxs-lookup"><span data-stu-id="554b9-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="554b9-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="554b9-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="554b9-141">1.6</span><span class="sxs-lookup"><span data-stu-id="554b9-141">1.6</span></span> |
|[<span data-ttu-id="554b9-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="554b9-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="554b9-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="554b9-143">ReadItem</span></span>|
|[<span data-ttu-id="554b9-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="554b9-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="554b9-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="554b9-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="554b9-146">Пример</span><span class="sxs-lookup"><span data-stu-id="554b9-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="554b9-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="554b9-147">displayName :String</span></span>

<span data-ttu-id="554b9-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="554b9-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="554b9-149">Тип</span><span class="sxs-lookup"><span data-stu-id="554b9-149">Type</span></span>

*   <span data-ttu-id="554b9-150">String</span><span class="sxs-lookup"><span data-stu-id="554b9-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="554b9-151">Требования</span><span class="sxs-lookup"><span data-stu-id="554b9-151">Requirements</span></span>

|<span data-ttu-id="554b9-152">Требование</span><span class="sxs-lookup"><span data-stu-id="554b9-152">Requirement</span></span>| <span data-ttu-id="554b9-153">Значение</span><span class="sxs-lookup"><span data-stu-id="554b9-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="554b9-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="554b9-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="554b9-155">1.0</span><span class="sxs-lookup"><span data-stu-id="554b9-155">1.0</span></span>|
|[<span data-ttu-id="554b9-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="554b9-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="554b9-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="554b9-157">ReadItem</span></span>|
|[<span data-ttu-id="554b9-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="554b9-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="554b9-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="554b9-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="554b9-160">Пример</span><span class="sxs-lookup"><span data-stu-id="554b9-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="554b9-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="554b9-161">emailAddress :String</span></span>

<span data-ttu-id="554b9-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="554b9-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="554b9-163">Тип</span><span class="sxs-lookup"><span data-stu-id="554b9-163">Type</span></span>

*   <span data-ttu-id="554b9-164">String</span><span class="sxs-lookup"><span data-stu-id="554b9-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="554b9-165">Требования</span><span class="sxs-lookup"><span data-stu-id="554b9-165">Requirements</span></span>

|<span data-ttu-id="554b9-166">Требование</span><span class="sxs-lookup"><span data-stu-id="554b9-166">Requirement</span></span>| <span data-ttu-id="554b9-167">Значение</span><span class="sxs-lookup"><span data-stu-id="554b9-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="554b9-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="554b9-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="554b9-169">1.0</span><span class="sxs-lookup"><span data-stu-id="554b9-169">1.0</span></span>|
|[<span data-ttu-id="554b9-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="554b9-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="554b9-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="554b9-171">ReadItem</span></span>|
|[<span data-ttu-id="554b9-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="554b9-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="554b9-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="554b9-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="554b9-174">Пример</span><span class="sxs-lookup"><span data-stu-id="554b9-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="554b9-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="554b9-175">timeZone :String</span></span>

<span data-ttu-id="554b9-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="554b9-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="554b9-177">Тип</span><span class="sxs-lookup"><span data-stu-id="554b9-177">Type</span></span>

*   <span data-ttu-id="554b9-178">String</span><span class="sxs-lookup"><span data-stu-id="554b9-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="554b9-179">Требования</span><span class="sxs-lookup"><span data-stu-id="554b9-179">Requirements</span></span>

|<span data-ttu-id="554b9-180">Требование</span><span class="sxs-lookup"><span data-stu-id="554b9-180">Requirement</span></span>| <span data-ttu-id="554b9-181">Значение</span><span class="sxs-lookup"><span data-stu-id="554b9-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="554b9-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="554b9-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="554b9-183">1.0</span><span class="sxs-lookup"><span data-stu-id="554b9-183">1.0</span></span>|
|[<span data-ttu-id="554b9-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="554b9-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="554b9-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="554b9-185">ReadItem</span></span>|
|[<span data-ttu-id="554b9-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="554b9-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="554b9-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="554b9-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="554b9-188">Пример</span><span class="sxs-lookup"><span data-stu-id="554b9-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
