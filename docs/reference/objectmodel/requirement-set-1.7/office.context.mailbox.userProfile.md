---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,7
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 335c20a242d08031e6f21c48e99c8527dab6d714
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871740"
---
# <a name="userprofile"></a><span data-ttu-id="467e5-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="467e5-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="467e5-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="467e5-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="467e5-104">Требования</span><span class="sxs-lookup"><span data-stu-id="467e5-104">Requirements</span></span>

|<span data-ttu-id="467e5-105">Требование</span><span class="sxs-lookup"><span data-stu-id="467e5-105">Requirement</span></span>| <span data-ttu-id="467e5-106">Значение</span><span class="sxs-lookup"><span data-stu-id="467e5-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="467e5-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="467e5-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="467e5-108">1.0</span><span class="sxs-lookup"><span data-stu-id="467e5-108">1.0</span></span>|
|[<span data-ttu-id="467e5-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="467e5-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="467e5-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="467e5-110">ReadItem</span></span>|
|[<span data-ttu-id="467e5-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="467e5-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="467e5-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="467e5-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="467e5-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="467e5-113">Members and methods</span></span>

| <span data-ttu-id="467e5-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="467e5-114">Member</span></span> | <span data-ttu-id="467e5-115">Тип</span><span class="sxs-lookup"><span data-stu-id="467e5-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="467e5-116">accountType</span><span class="sxs-lookup"><span data-stu-id="467e5-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="467e5-117">Member</span><span class="sxs-lookup"><span data-stu-id="467e5-117">Member</span></span> |
| [<span data-ttu-id="467e5-118">displayName</span><span class="sxs-lookup"><span data-stu-id="467e5-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="467e5-119">Member</span><span class="sxs-lookup"><span data-stu-id="467e5-119">Member</span></span> |
| [<span data-ttu-id="467e5-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="467e5-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="467e5-121">Member</span><span class="sxs-lookup"><span data-stu-id="467e5-121">Member</span></span> |
| [<span data-ttu-id="467e5-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="467e5-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="467e5-123">Member</span><span class="sxs-lookup"><span data-stu-id="467e5-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="467e5-124">Members</span><span class="sxs-lookup"><span data-stu-id="467e5-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="467e5-125">accountType: строка</span><span class="sxs-lookup"><span data-stu-id="467e5-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="467e5-126">В настоящее время этот элемент поддерживается только Outlook 2016 для Mac (сборка 16.9.1212 или более поздняя).</span><span class="sxs-lookup"><span data-stu-id="467e5-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="467e5-127">Возвращает тип учетной записи пользователя, связанного с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="467e5-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="467e5-128">Возможные значения перечислены в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="467e5-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="467e5-129">Значение</span><span class="sxs-lookup"><span data-stu-id="467e5-129">Value</span></span> | <span data-ttu-id="467e5-130">Описание</span><span class="sxs-lookup"><span data-stu-id="467e5-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="467e5-131">Почтовый ящик находится на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="467e5-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="467e5-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="467e5-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="467e5-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="467e5-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="467e5-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="467e5-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="467e5-135">Тип</span><span class="sxs-lookup"><span data-stu-id="467e5-135">Type</span></span>

*   <span data-ttu-id="467e5-136">String</span><span class="sxs-lookup"><span data-stu-id="467e5-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="467e5-137">Требования</span><span class="sxs-lookup"><span data-stu-id="467e5-137">Requirements</span></span>

|<span data-ttu-id="467e5-138">Требование</span><span class="sxs-lookup"><span data-stu-id="467e5-138">Requirement</span></span>| <span data-ttu-id="467e5-139">Значение</span><span class="sxs-lookup"><span data-stu-id="467e5-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="467e5-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="467e5-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="467e5-141">1.6</span><span class="sxs-lookup"><span data-stu-id="467e5-141">1.6</span></span> |
|[<span data-ttu-id="467e5-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="467e5-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="467e5-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="467e5-143">ReadItem</span></span>|
|[<span data-ttu-id="467e5-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="467e5-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="467e5-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="467e5-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="467e5-146">Пример</span><span class="sxs-lookup"><span data-stu-id="467e5-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="467e5-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="467e5-147">displayName :String</span></span>

<span data-ttu-id="467e5-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="467e5-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="467e5-149">Тип</span><span class="sxs-lookup"><span data-stu-id="467e5-149">Type</span></span>

*   <span data-ttu-id="467e5-150">String</span><span class="sxs-lookup"><span data-stu-id="467e5-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="467e5-151">Требования</span><span class="sxs-lookup"><span data-stu-id="467e5-151">Requirements</span></span>

|<span data-ttu-id="467e5-152">Требование</span><span class="sxs-lookup"><span data-stu-id="467e5-152">Requirement</span></span>| <span data-ttu-id="467e5-153">Значение</span><span class="sxs-lookup"><span data-stu-id="467e5-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="467e5-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="467e5-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="467e5-155">1.0</span><span class="sxs-lookup"><span data-stu-id="467e5-155">1.0</span></span>|
|[<span data-ttu-id="467e5-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="467e5-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="467e5-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="467e5-157">ReadItem</span></span>|
|[<span data-ttu-id="467e5-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="467e5-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="467e5-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="467e5-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="467e5-160">Пример</span><span class="sxs-lookup"><span data-stu-id="467e5-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="467e5-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="467e5-161">emailAddress :String</span></span>

<span data-ttu-id="467e5-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="467e5-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="467e5-163">Тип</span><span class="sxs-lookup"><span data-stu-id="467e5-163">Type</span></span>

*   <span data-ttu-id="467e5-164">String</span><span class="sxs-lookup"><span data-stu-id="467e5-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="467e5-165">Требования</span><span class="sxs-lookup"><span data-stu-id="467e5-165">Requirements</span></span>

|<span data-ttu-id="467e5-166">Требование</span><span class="sxs-lookup"><span data-stu-id="467e5-166">Requirement</span></span>| <span data-ttu-id="467e5-167">Значение</span><span class="sxs-lookup"><span data-stu-id="467e5-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="467e5-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="467e5-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="467e5-169">1.0</span><span class="sxs-lookup"><span data-stu-id="467e5-169">1.0</span></span>|
|[<span data-ttu-id="467e5-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="467e5-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="467e5-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="467e5-171">ReadItem</span></span>|
|[<span data-ttu-id="467e5-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="467e5-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="467e5-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="467e5-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="467e5-174">Пример</span><span class="sxs-lookup"><span data-stu-id="467e5-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="467e5-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="467e5-175">timeZone :String</span></span>

<span data-ttu-id="467e5-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="467e5-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="467e5-177">Тип</span><span class="sxs-lookup"><span data-stu-id="467e5-177">Type</span></span>

*   <span data-ttu-id="467e5-178">String</span><span class="sxs-lookup"><span data-stu-id="467e5-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="467e5-179">Требования</span><span class="sxs-lookup"><span data-stu-id="467e5-179">Requirements</span></span>

|<span data-ttu-id="467e5-180">Требование</span><span class="sxs-lookup"><span data-stu-id="467e5-180">Requirement</span></span>| <span data-ttu-id="467e5-181">Значение</span><span class="sxs-lookup"><span data-stu-id="467e5-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="467e5-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="467e5-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="467e5-183">1.0</span><span class="sxs-lookup"><span data-stu-id="467e5-183">1.0</span></span>|
|[<span data-ttu-id="467e5-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="467e5-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="467e5-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="467e5-185">ReadItem</span></span>|
|[<span data-ttu-id="467e5-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="467e5-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="467e5-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="467e5-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="467e5-188">Пример</span><span class="sxs-lookup"><span data-stu-id="467e5-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
