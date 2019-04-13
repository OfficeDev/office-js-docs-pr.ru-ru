---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 8cfee874bbb5183d62cc3a9ce8b042a76617ec72
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838524"
---
# <a name="userprofile"></a><span data-ttu-id="59b5c-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="59b5c-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="59b5c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="59b5c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="59b5c-104">Требования</span><span class="sxs-lookup"><span data-stu-id="59b5c-104">Requirements</span></span>

|<span data-ttu-id="59b5c-105">Требование</span><span class="sxs-lookup"><span data-stu-id="59b5c-105">Requirement</span></span>| <span data-ttu-id="59b5c-106">Значение</span><span class="sxs-lookup"><span data-stu-id="59b5c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="59b5c-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="59b5c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="59b5c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="59b5c-108">1.0</span></span>|
|[<span data-ttu-id="59b5c-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="59b5c-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="59b5c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59b5c-110">ReadItem</span></span>|
|[<span data-ttu-id="59b5c-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="59b5c-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="59b5c-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="59b5c-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="59b5c-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="59b5c-113">Members and methods</span></span>

| <span data-ttu-id="59b5c-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="59b5c-114">Member</span></span> | <span data-ttu-id="59b5c-115">Тип</span><span class="sxs-lookup"><span data-stu-id="59b5c-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="59b5c-116">accountType</span><span class="sxs-lookup"><span data-stu-id="59b5c-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="59b5c-117">Member</span><span class="sxs-lookup"><span data-stu-id="59b5c-117">Member</span></span> |
| [<span data-ttu-id="59b5c-118">displayName</span><span class="sxs-lookup"><span data-stu-id="59b5c-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="59b5c-119">Member</span><span class="sxs-lookup"><span data-stu-id="59b5c-119">Member</span></span> |
| [<span data-ttu-id="59b5c-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="59b5c-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="59b5c-121">Member</span><span class="sxs-lookup"><span data-stu-id="59b5c-121">Member</span></span> |
| [<span data-ttu-id="59b5c-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="59b5c-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="59b5c-123">Member</span><span class="sxs-lookup"><span data-stu-id="59b5c-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="59b5c-124">Members</span><span class="sxs-lookup"><span data-stu-id="59b5c-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="59b5c-125">accountType: строка</span><span class="sxs-lookup"><span data-stu-id="59b5c-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="59b5c-126">В настоящее время этот элемент поддерживается только Outlook 2016 для Mac (сборка 16.9.1212 или более поздняя).</span><span class="sxs-lookup"><span data-stu-id="59b5c-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="59b5c-127">Возвращает тип учетной записи пользователя, связанного с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="59b5c-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="59b5c-128">Возможные значения перечислены в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="59b5c-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="59b5c-129">Значение</span><span class="sxs-lookup"><span data-stu-id="59b5c-129">Value</span></span> | <span data-ttu-id="59b5c-130">Описание</span><span class="sxs-lookup"><span data-stu-id="59b5c-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="59b5c-131">Почтовый ящик находится на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="59b5c-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="59b5c-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="59b5c-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="59b5c-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="59b5c-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="59b5c-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="59b5c-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="59b5c-135">Тип</span><span class="sxs-lookup"><span data-stu-id="59b5c-135">Type</span></span>

*   <span data-ttu-id="59b5c-136">String</span><span class="sxs-lookup"><span data-stu-id="59b5c-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59b5c-137">Требования</span><span class="sxs-lookup"><span data-stu-id="59b5c-137">Requirements</span></span>

|<span data-ttu-id="59b5c-138">Требование</span><span class="sxs-lookup"><span data-stu-id="59b5c-138">Requirement</span></span>| <span data-ttu-id="59b5c-139">Значение</span><span class="sxs-lookup"><span data-stu-id="59b5c-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="59b5c-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="59b5c-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="59b5c-141">1.6</span><span class="sxs-lookup"><span data-stu-id="59b5c-141">1.6</span></span> |
|[<span data-ttu-id="59b5c-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="59b5c-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="59b5c-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59b5c-143">ReadItem</span></span>|
|[<span data-ttu-id="59b5c-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="59b5c-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="59b5c-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="59b5c-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59b5c-146">Пример</span><span class="sxs-lookup"><span data-stu-id="59b5c-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

####  <a name="displayname-string"></a><span data-ttu-id="59b5c-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="59b5c-147">displayName :String</span></span>

<span data-ttu-id="59b5c-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="59b5c-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="59b5c-149">Тип</span><span class="sxs-lookup"><span data-stu-id="59b5c-149">Type</span></span>

*   <span data-ttu-id="59b5c-150">String</span><span class="sxs-lookup"><span data-stu-id="59b5c-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59b5c-151">Требования</span><span class="sxs-lookup"><span data-stu-id="59b5c-151">Requirements</span></span>

|<span data-ttu-id="59b5c-152">Требование</span><span class="sxs-lookup"><span data-stu-id="59b5c-152">Requirement</span></span>| <span data-ttu-id="59b5c-153">Значение</span><span class="sxs-lookup"><span data-stu-id="59b5c-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="59b5c-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="59b5c-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="59b5c-155">1.0</span><span class="sxs-lookup"><span data-stu-id="59b5c-155">1.0</span></span>|
|[<span data-ttu-id="59b5c-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="59b5c-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="59b5c-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59b5c-157">ReadItem</span></span>|
|[<span data-ttu-id="59b5c-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="59b5c-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="59b5c-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="59b5c-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59b5c-160">Пример</span><span class="sxs-lookup"><span data-stu-id="59b5c-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

####  <a name="emailaddress-string"></a><span data-ttu-id="59b5c-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="59b5c-161">emailAddress :String</span></span>

<span data-ttu-id="59b5c-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="59b5c-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="59b5c-163">Тип</span><span class="sxs-lookup"><span data-stu-id="59b5c-163">Type</span></span>

*   <span data-ttu-id="59b5c-164">String</span><span class="sxs-lookup"><span data-stu-id="59b5c-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59b5c-165">Требования</span><span class="sxs-lookup"><span data-stu-id="59b5c-165">Requirements</span></span>

|<span data-ttu-id="59b5c-166">Требование</span><span class="sxs-lookup"><span data-stu-id="59b5c-166">Requirement</span></span>| <span data-ttu-id="59b5c-167">Значение</span><span class="sxs-lookup"><span data-stu-id="59b5c-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="59b5c-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="59b5c-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="59b5c-169">1.0</span><span class="sxs-lookup"><span data-stu-id="59b5c-169">1.0</span></span>|
|[<span data-ttu-id="59b5c-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="59b5c-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="59b5c-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59b5c-171">ReadItem</span></span>|
|[<span data-ttu-id="59b5c-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="59b5c-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="59b5c-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="59b5c-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59b5c-174">Пример</span><span class="sxs-lookup"><span data-stu-id="59b5c-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

####  <a name="timezone-string"></a><span data-ttu-id="59b5c-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="59b5c-175">timeZone :String</span></span>

<span data-ttu-id="59b5c-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="59b5c-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="59b5c-177">Тип</span><span class="sxs-lookup"><span data-stu-id="59b5c-177">Type</span></span>

*   <span data-ttu-id="59b5c-178">String</span><span class="sxs-lookup"><span data-stu-id="59b5c-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59b5c-179">Требования</span><span class="sxs-lookup"><span data-stu-id="59b5c-179">Requirements</span></span>

|<span data-ttu-id="59b5c-180">Требование</span><span class="sxs-lookup"><span data-stu-id="59b5c-180">Requirement</span></span>| <span data-ttu-id="59b5c-181">Значение</span><span class="sxs-lookup"><span data-stu-id="59b5c-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="59b5c-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="59b5c-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="59b5c-183">1.0</span><span class="sxs-lookup"><span data-stu-id="59b5c-183">1.0</span></span>|
|[<span data-ttu-id="59b5c-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="59b5c-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="59b5c-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59b5c-185">ReadItem</span></span>|
|[<span data-ttu-id="59b5c-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="59b5c-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="59b5c-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="59b5c-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59b5c-188">Пример</span><span class="sxs-lookup"><span data-stu-id="59b5c-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
