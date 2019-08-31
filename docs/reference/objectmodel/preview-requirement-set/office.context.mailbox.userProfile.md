---
title: Office. Context. Mailbox. userProfile — Предварительная версия набора требований
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 5941c4e1276535091a3ffcf5b2fb6aa972ed8c4d
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696472"
---
# <a name="userprofile"></a><span data-ttu-id="035e4-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="035e4-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="035e4-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="035e4-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="035e4-104">Требования</span><span class="sxs-lookup"><span data-stu-id="035e4-104">Requirements</span></span>

|<span data-ttu-id="035e4-105">Требование</span><span class="sxs-lookup"><span data-stu-id="035e4-105">Requirement</span></span>| <span data-ttu-id="035e4-106">Значение</span><span class="sxs-lookup"><span data-stu-id="035e4-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="035e4-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="035e4-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="035e4-108">1.0</span><span class="sxs-lookup"><span data-stu-id="035e4-108">1.0</span></span>|
|[<span data-ttu-id="035e4-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="035e4-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="035e4-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="035e4-110">ReadItem</span></span>|
|[<span data-ttu-id="035e4-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="035e4-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="035e4-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="035e4-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="035e4-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="035e4-113">Members and methods</span></span>

| <span data-ttu-id="035e4-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="035e4-114">Member</span></span> | <span data-ttu-id="035e4-115">Тип</span><span class="sxs-lookup"><span data-stu-id="035e4-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="035e4-116">accountType</span><span class="sxs-lookup"><span data-stu-id="035e4-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="035e4-117">Member</span><span class="sxs-lookup"><span data-stu-id="035e4-117">Member</span></span> |
| [<span data-ttu-id="035e4-118">displayName</span><span class="sxs-lookup"><span data-stu-id="035e4-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="035e4-119">Member</span><span class="sxs-lookup"><span data-stu-id="035e4-119">Member</span></span> |
| [<span data-ttu-id="035e4-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="035e4-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="035e4-121">Member</span><span class="sxs-lookup"><span data-stu-id="035e4-121">Member</span></span> |
| [<span data-ttu-id="035e4-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="035e4-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="035e4-123">Member</span><span class="sxs-lookup"><span data-stu-id="035e4-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="035e4-124">Members</span><span class="sxs-lookup"><span data-stu-id="035e4-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="035e4-125">accountType: строка</span><span class="sxs-lookup"><span data-stu-id="035e4-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="035e4-126">В настоящее время этот элемент поддерживается только в Outlook 2016 или более поздней версии в Mac (сборка 16.9.1212 или более поздняя).</span><span class="sxs-lookup"><span data-stu-id="035e4-126">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="035e4-127">Возвращает тип учетной записи пользователя, связанного с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="035e4-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="035e4-128">Возможные значения перечислены в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="035e4-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="035e4-129">Значение</span><span class="sxs-lookup"><span data-stu-id="035e4-129">Value</span></span> | <span data-ttu-id="035e4-130">Описание</span><span class="sxs-lookup"><span data-stu-id="035e4-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="035e4-131">Почтовый ящик находится на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="035e4-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="035e4-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="035e4-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="035e4-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="035e4-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="035e4-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="035e4-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="035e4-135">Тип</span><span class="sxs-lookup"><span data-stu-id="035e4-135">Type</span></span>

*   <span data-ttu-id="035e4-136">String</span><span class="sxs-lookup"><span data-stu-id="035e4-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="035e4-137">Требования</span><span class="sxs-lookup"><span data-stu-id="035e4-137">Requirements</span></span>

|<span data-ttu-id="035e4-138">Требование</span><span class="sxs-lookup"><span data-stu-id="035e4-138">Requirement</span></span>| <span data-ttu-id="035e4-139">Значение</span><span class="sxs-lookup"><span data-stu-id="035e4-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="035e4-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="035e4-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="035e4-141">1.6</span><span class="sxs-lookup"><span data-stu-id="035e4-141">1.6</span></span> |
|[<span data-ttu-id="035e4-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="035e4-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="035e4-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="035e4-143">ReadItem</span></span>|
|[<span data-ttu-id="035e4-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="035e4-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="035e4-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="035e4-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="035e4-146">Пример</span><span class="sxs-lookup"><span data-stu-id="035e4-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="035e4-147">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="035e4-147">displayName: String</span></span>

<span data-ttu-id="035e4-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="035e4-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="035e4-149">Тип</span><span class="sxs-lookup"><span data-stu-id="035e4-149">Type</span></span>

*   <span data-ttu-id="035e4-150">String</span><span class="sxs-lookup"><span data-stu-id="035e4-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="035e4-151">Требования</span><span class="sxs-lookup"><span data-stu-id="035e4-151">Requirements</span></span>

|<span data-ttu-id="035e4-152">Требование</span><span class="sxs-lookup"><span data-stu-id="035e4-152">Requirement</span></span>| <span data-ttu-id="035e4-153">Значение</span><span class="sxs-lookup"><span data-stu-id="035e4-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="035e4-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="035e4-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="035e4-155">1.0</span><span class="sxs-lookup"><span data-stu-id="035e4-155">1.0</span></span>|
|[<span data-ttu-id="035e4-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="035e4-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="035e4-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="035e4-157">ReadItem</span></span>|
|[<span data-ttu-id="035e4-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="035e4-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="035e4-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="035e4-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="035e4-160">Пример</span><span class="sxs-lookup"><span data-stu-id="035e4-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="035e4-161">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="035e4-161">emailAddress: String</span></span>

<span data-ttu-id="035e4-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="035e4-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="035e4-163">Тип</span><span class="sxs-lookup"><span data-stu-id="035e4-163">Type</span></span>

*   <span data-ttu-id="035e4-164">String</span><span class="sxs-lookup"><span data-stu-id="035e4-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="035e4-165">Требования</span><span class="sxs-lookup"><span data-stu-id="035e4-165">Requirements</span></span>

|<span data-ttu-id="035e4-166">Требование</span><span class="sxs-lookup"><span data-stu-id="035e4-166">Requirement</span></span>| <span data-ttu-id="035e4-167">Значение</span><span class="sxs-lookup"><span data-stu-id="035e4-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="035e4-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="035e4-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="035e4-169">1.0</span><span class="sxs-lookup"><span data-stu-id="035e4-169">1.0</span></span>|
|[<span data-ttu-id="035e4-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="035e4-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="035e4-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="035e4-171">ReadItem</span></span>|
|[<span data-ttu-id="035e4-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="035e4-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="035e4-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="035e4-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="035e4-174">Пример</span><span class="sxs-lookup"><span data-stu-id="035e4-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="035e4-175">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="035e4-175">timeZone: String</span></span>

<span data-ttu-id="035e4-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="035e4-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="035e4-177">Тип</span><span class="sxs-lookup"><span data-stu-id="035e4-177">Type</span></span>

*   <span data-ttu-id="035e4-178">String</span><span class="sxs-lookup"><span data-stu-id="035e4-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="035e4-179">Требования</span><span class="sxs-lookup"><span data-stu-id="035e4-179">Requirements</span></span>

|<span data-ttu-id="035e4-180">Требование</span><span class="sxs-lookup"><span data-stu-id="035e4-180">Requirement</span></span>| <span data-ttu-id="035e4-181">Значение</span><span class="sxs-lookup"><span data-stu-id="035e4-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="035e4-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="035e4-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="035e4-183">1.0</span><span class="sxs-lookup"><span data-stu-id="035e4-183">1.0</span></span>|
|[<span data-ttu-id="035e4-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="035e4-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="035e4-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="035e4-185">ReadItem</span></span>|
|[<span data-ttu-id="035e4-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="035e4-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="035e4-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="035e4-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="035e4-188">Пример</span><span class="sxs-lookup"><span data-stu-id="035e4-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
