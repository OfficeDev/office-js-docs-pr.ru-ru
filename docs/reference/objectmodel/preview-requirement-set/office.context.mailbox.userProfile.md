---
title: Office. Context. Mailbox. userProfile — Предварительная версия набора требований
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 29111314f16bb9c6518b350254a3036ffa125796
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451964"
---
# <a name="userprofile"></a><span data-ttu-id="cbb87-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="cbb87-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="cbb87-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="cbb87-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="cbb87-104">Требования</span><span class="sxs-lookup"><span data-stu-id="cbb87-104">Requirements</span></span>

|<span data-ttu-id="cbb87-105">Требование</span><span class="sxs-lookup"><span data-stu-id="cbb87-105">Requirement</span></span>| <span data-ttu-id="cbb87-106">Значение</span><span class="sxs-lookup"><span data-stu-id="cbb87-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="cbb87-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cbb87-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cbb87-108">1.0</span><span class="sxs-lookup"><span data-stu-id="cbb87-108">1.0</span></span>|
|[<span data-ttu-id="cbb87-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cbb87-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cbb87-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cbb87-110">ReadItem</span></span>|
|[<span data-ttu-id="cbb87-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cbb87-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cbb87-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cbb87-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="cbb87-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="cbb87-113">Members and methods</span></span>

| <span data-ttu-id="cbb87-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="cbb87-114">Member</span></span> | <span data-ttu-id="cbb87-115">Тип</span><span class="sxs-lookup"><span data-stu-id="cbb87-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="cbb87-116">accountType</span><span class="sxs-lookup"><span data-stu-id="cbb87-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="cbb87-117">Member</span><span class="sxs-lookup"><span data-stu-id="cbb87-117">Member</span></span> |
| [<span data-ttu-id="cbb87-118">displayName</span><span class="sxs-lookup"><span data-stu-id="cbb87-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="cbb87-119">Member</span><span class="sxs-lookup"><span data-stu-id="cbb87-119">Member</span></span> |
| [<span data-ttu-id="cbb87-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="cbb87-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="cbb87-121">Member</span><span class="sxs-lookup"><span data-stu-id="cbb87-121">Member</span></span> |
| [<span data-ttu-id="cbb87-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="cbb87-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="cbb87-123">Member</span><span class="sxs-lookup"><span data-stu-id="cbb87-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="cbb87-124">Members</span><span class="sxs-lookup"><span data-stu-id="cbb87-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="cbb87-125">accountType: строка</span><span class="sxs-lookup"><span data-stu-id="cbb87-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="cbb87-126">В настоящее время этот элемент поддерживается только в Outlook 2016 или более поздней версии для Mac (сборка 16.9.1212 или более поздняя).</span><span class="sxs-lookup"><span data-stu-id="cbb87-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="cbb87-127">Возвращает тип учетной записи пользователя, связанного с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="cbb87-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="cbb87-128">Возможные значения перечислены в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="cbb87-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="cbb87-129">Значение</span><span class="sxs-lookup"><span data-stu-id="cbb87-129">Value</span></span> | <span data-ttu-id="cbb87-130">Описание</span><span class="sxs-lookup"><span data-stu-id="cbb87-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="cbb87-131">Почтовый ящик находится на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="cbb87-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="cbb87-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="cbb87-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="cbb87-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="cbb87-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="cbb87-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="cbb87-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="cbb87-135">Тип</span><span class="sxs-lookup"><span data-stu-id="cbb87-135">Type</span></span>

*   <span data-ttu-id="cbb87-136">String</span><span class="sxs-lookup"><span data-stu-id="cbb87-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cbb87-137">Требования</span><span class="sxs-lookup"><span data-stu-id="cbb87-137">Requirements</span></span>

|<span data-ttu-id="cbb87-138">Требование</span><span class="sxs-lookup"><span data-stu-id="cbb87-138">Requirement</span></span>| <span data-ttu-id="cbb87-139">Значение</span><span class="sxs-lookup"><span data-stu-id="cbb87-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="cbb87-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="cbb87-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cbb87-141">1.6</span><span class="sxs-lookup"><span data-stu-id="cbb87-141">1.6</span></span> |
|[<span data-ttu-id="cbb87-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cbb87-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cbb87-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cbb87-143">ReadItem</span></span>|
|[<span data-ttu-id="cbb87-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cbb87-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cbb87-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cbb87-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cbb87-146">Пример</span><span class="sxs-lookup"><span data-stu-id="cbb87-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

####  <a name="displayname-string"></a><span data-ttu-id="cbb87-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="cbb87-147">displayName :String</span></span>

<span data-ttu-id="cbb87-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="cbb87-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="cbb87-149">Тип</span><span class="sxs-lookup"><span data-stu-id="cbb87-149">Type</span></span>

*   <span data-ttu-id="cbb87-150">String</span><span class="sxs-lookup"><span data-stu-id="cbb87-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cbb87-151">Требования</span><span class="sxs-lookup"><span data-stu-id="cbb87-151">Requirements</span></span>

|<span data-ttu-id="cbb87-152">Требование</span><span class="sxs-lookup"><span data-stu-id="cbb87-152">Requirement</span></span>| <span data-ttu-id="cbb87-153">Значение</span><span class="sxs-lookup"><span data-stu-id="cbb87-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="cbb87-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cbb87-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cbb87-155">1.0</span><span class="sxs-lookup"><span data-stu-id="cbb87-155">1.0</span></span>|
|[<span data-ttu-id="cbb87-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cbb87-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cbb87-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cbb87-157">ReadItem</span></span>|
|[<span data-ttu-id="cbb87-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cbb87-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cbb87-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cbb87-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cbb87-160">Пример</span><span class="sxs-lookup"><span data-stu-id="cbb87-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

####  <a name="emailaddress-string"></a><span data-ttu-id="cbb87-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="cbb87-161">emailAddress :String</span></span>

<span data-ttu-id="cbb87-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="cbb87-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="cbb87-163">Тип</span><span class="sxs-lookup"><span data-stu-id="cbb87-163">Type</span></span>

*   <span data-ttu-id="cbb87-164">String</span><span class="sxs-lookup"><span data-stu-id="cbb87-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cbb87-165">Требования</span><span class="sxs-lookup"><span data-stu-id="cbb87-165">Requirements</span></span>

|<span data-ttu-id="cbb87-166">Требование</span><span class="sxs-lookup"><span data-stu-id="cbb87-166">Requirement</span></span>| <span data-ttu-id="cbb87-167">Значение</span><span class="sxs-lookup"><span data-stu-id="cbb87-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="cbb87-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cbb87-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cbb87-169">1.0</span><span class="sxs-lookup"><span data-stu-id="cbb87-169">1.0</span></span>|
|[<span data-ttu-id="cbb87-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cbb87-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cbb87-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cbb87-171">ReadItem</span></span>|
|[<span data-ttu-id="cbb87-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cbb87-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cbb87-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cbb87-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cbb87-174">Пример</span><span class="sxs-lookup"><span data-stu-id="cbb87-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

####  <a name="timezone-string"></a><span data-ttu-id="cbb87-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="cbb87-175">timeZone :String</span></span>

<span data-ttu-id="cbb87-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="cbb87-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="cbb87-177">Тип</span><span class="sxs-lookup"><span data-stu-id="cbb87-177">Type</span></span>

*   <span data-ttu-id="cbb87-178">String</span><span class="sxs-lookup"><span data-stu-id="cbb87-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cbb87-179">Требования</span><span class="sxs-lookup"><span data-stu-id="cbb87-179">Requirements</span></span>

|<span data-ttu-id="cbb87-180">Требование</span><span class="sxs-lookup"><span data-stu-id="cbb87-180">Requirement</span></span>| <span data-ttu-id="cbb87-181">Значение</span><span class="sxs-lookup"><span data-stu-id="cbb87-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="cbb87-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cbb87-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cbb87-183">1.0</span><span class="sxs-lookup"><span data-stu-id="cbb87-183">1.0</span></span>|
|[<span data-ttu-id="cbb87-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cbb87-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cbb87-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cbb87-185">ReadItem</span></span>|
|[<span data-ttu-id="cbb87-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cbb87-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cbb87-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cbb87-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cbb87-188">Пример</span><span class="sxs-lookup"><span data-stu-id="cbb87-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
