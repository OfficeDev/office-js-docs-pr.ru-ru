---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,7
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 036f18e4cb98cfe510a19d85a5a79f393ca8bd17
ms.sourcegitcommit: 7462409209264dc7f8f89f3808a7a6249fcd739e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/26/2019
ms.locfileid: "33353295"
---
# <a name="userprofile"></a><span data-ttu-id="6b239-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="6b239-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="6b239-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="6b239-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b239-104">Требования</span><span class="sxs-lookup"><span data-stu-id="6b239-104">Requirements</span></span>

|<span data-ttu-id="6b239-105">Требование</span><span class="sxs-lookup"><span data-stu-id="6b239-105">Requirement</span></span>| <span data-ttu-id="6b239-106">Значение</span><span class="sxs-lookup"><span data-stu-id="6b239-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b239-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6b239-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b239-108">1.0</span><span class="sxs-lookup"><span data-stu-id="6b239-108">1.0</span></span>|
|[<span data-ttu-id="6b239-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6b239-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b239-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b239-110">ReadItem</span></span>|
|[<span data-ttu-id="6b239-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6b239-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b239-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6b239-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6b239-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="6b239-113">Members and methods</span></span>

| <span data-ttu-id="6b239-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="6b239-114">Member</span></span> | <span data-ttu-id="6b239-115">Тип</span><span class="sxs-lookup"><span data-stu-id="6b239-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6b239-116">accountType</span><span class="sxs-lookup"><span data-stu-id="6b239-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="6b239-117">Member</span><span class="sxs-lookup"><span data-stu-id="6b239-117">Member</span></span> |
| [<span data-ttu-id="6b239-118">displayName</span><span class="sxs-lookup"><span data-stu-id="6b239-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="6b239-119">Member</span><span class="sxs-lookup"><span data-stu-id="6b239-119">Member</span></span> |
| [<span data-ttu-id="6b239-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="6b239-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="6b239-121">Member</span><span class="sxs-lookup"><span data-stu-id="6b239-121">Member</span></span> |
| [<span data-ttu-id="6b239-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="6b239-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="6b239-123">Member</span><span class="sxs-lookup"><span data-stu-id="6b239-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="6b239-124">Members</span><span class="sxs-lookup"><span data-stu-id="6b239-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="6b239-125">accountType: строка</span><span class="sxs-lookup"><span data-stu-id="6b239-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="6b239-126">В настоящее время этот элемент поддерживается только Outlook 2016 для Mac (сборка 16.9.1212 или более поздняя).</span><span class="sxs-lookup"><span data-stu-id="6b239-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="6b239-127">Возвращает тип учетной записи пользователя, связанного с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="6b239-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="6b239-128">Возможные значения перечислены в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="6b239-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="6b239-129">Значение</span><span class="sxs-lookup"><span data-stu-id="6b239-129">Value</span></span> | <span data-ttu-id="6b239-130">Описание</span><span class="sxs-lookup"><span data-stu-id="6b239-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="6b239-131">Почтовый ящик находится на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="6b239-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="6b239-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="6b239-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="6b239-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="6b239-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="6b239-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="6b239-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="6b239-135">Тип</span><span class="sxs-lookup"><span data-stu-id="6b239-135">Type</span></span>

*   <span data-ttu-id="6b239-136">String</span><span class="sxs-lookup"><span data-stu-id="6b239-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b239-137">Требования</span><span class="sxs-lookup"><span data-stu-id="6b239-137">Requirements</span></span>

|<span data-ttu-id="6b239-138">Требование</span><span class="sxs-lookup"><span data-stu-id="6b239-138">Requirement</span></span>| <span data-ttu-id="6b239-139">Значение</span><span class="sxs-lookup"><span data-stu-id="6b239-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b239-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="6b239-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b239-141">1.6</span><span class="sxs-lookup"><span data-stu-id="6b239-141">1.6</span></span> |
|[<span data-ttu-id="6b239-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6b239-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b239-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b239-143">ReadItem</span></span>|
|[<span data-ttu-id="6b239-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6b239-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b239-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6b239-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6b239-146">Пример</span><span class="sxs-lookup"><span data-stu-id="6b239-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

#### <a name="displayname-string"></a><span data-ttu-id="6b239-147">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="6b239-147">displayName: String</span></span>

<span data-ttu-id="6b239-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="6b239-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="6b239-149">Тип</span><span class="sxs-lookup"><span data-stu-id="6b239-149">Type</span></span>

*   <span data-ttu-id="6b239-150">String</span><span class="sxs-lookup"><span data-stu-id="6b239-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b239-151">Требования</span><span class="sxs-lookup"><span data-stu-id="6b239-151">Requirements</span></span>

|<span data-ttu-id="6b239-152">Требование</span><span class="sxs-lookup"><span data-stu-id="6b239-152">Requirement</span></span>| <span data-ttu-id="6b239-153">Значение</span><span class="sxs-lookup"><span data-stu-id="6b239-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b239-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6b239-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b239-155">1.0</span><span class="sxs-lookup"><span data-stu-id="6b239-155">1.0</span></span>|
|[<span data-ttu-id="6b239-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6b239-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b239-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b239-157">ReadItem</span></span>|
|[<span data-ttu-id="6b239-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6b239-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b239-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6b239-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6b239-160">Пример</span><span class="sxs-lookup"><span data-stu-id="6b239-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="6b239-161">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="6b239-161">emailAddress: String</span></span>

<span data-ttu-id="6b239-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="6b239-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="6b239-163">Тип</span><span class="sxs-lookup"><span data-stu-id="6b239-163">Type</span></span>

*   <span data-ttu-id="6b239-164">String</span><span class="sxs-lookup"><span data-stu-id="6b239-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b239-165">Требования</span><span class="sxs-lookup"><span data-stu-id="6b239-165">Requirements</span></span>

|<span data-ttu-id="6b239-166">Требование</span><span class="sxs-lookup"><span data-stu-id="6b239-166">Requirement</span></span>| <span data-ttu-id="6b239-167">Значение</span><span class="sxs-lookup"><span data-stu-id="6b239-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b239-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6b239-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b239-169">1.0</span><span class="sxs-lookup"><span data-stu-id="6b239-169">1.0</span></span>|
|[<span data-ttu-id="6b239-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6b239-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b239-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b239-171">ReadItem</span></span>|
|[<span data-ttu-id="6b239-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6b239-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b239-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6b239-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6b239-174">Пример</span><span class="sxs-lookup"><span data-stu-id="6b239-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

#### <a name="timezone-string"></a><span data-ttu-id="6b239-175">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="6b239-175">timeZone: String</span></span>

<span data-ttu-id="6b239-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="6b239-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="6b239-177">Тип</span><span class="sxs-lookup"><span data-stu-id="6b239-177">Type</span></span>

*   <span data-ttu-id="6b239-178">String</span><span class="sxs-lookup"><span data-stu-id="6b239-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b239-179">Требования</span><span class="sxs-lookup"><span data-stu-id="6b239-179">Requirements</span></span>

|<span data-ttu-id="6b239-180">Требование</span><span class="sxs-lookup"><span data-stu-id="6b239-180">Requirement</span></span>| <span data-ttu-id="6b239-181">Значение</span><span class="sxs-lookup"><span data-stu-id="6b239-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b239-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6b239-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b239-183">1.0</span><span class="sxs-lookup"><span data-stu-id="6b239-183">1.0</span></span>|
|[<span data-ttu-id="6b239-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6b239-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b239-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b239-185">ReadItem</span></span>|
|[<span data-ttu-id="6b239-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6b239-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="6b239-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6b239-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6b239-188">Пример</span><span class="sxs-lookup"><span data-stu-id="6b239-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
