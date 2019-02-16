---
title: Office.context.mailbox.userProfile — предварительная версия набора обязательных элементов
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 214434c988c01ecb1aef93f4067cd95bfe768ae9
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068177"
---
# <a name="userprofile"></a><span data-ttu-id="579d9-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="579d9-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="579d9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="579d9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="579d9-104">Требования</span><span class="sxs-lookup"><span data-stu-id="579d9-104">Requirements</span></span>

|<span data-ttu-id="579d9-105">Требование</span><span class="sxs-lookup"><span data-stu-id="579d9-105">Requirement</span></span>| <span data-ttu-id="579d9-106">Значение</span><span class="sxs-lookup"><span data-stu-id="579d9-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="579d9-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="579d9-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="579d9-108">1.0</span><span class="sxs-lookup"><span data-stu-id="579d9-108">1.0</span></span>|
|[<span data-ttu-id="579d9-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="579d9-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="579d9-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="579d9-110">ReadItem</span></span>|
|[<span data-ttu-id="579d9-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="579d9-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="579d9-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="579d9-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="579d9-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="579d9-113">Members and methods</span></span>

| <span data-ttu-id="579d9-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="579d9-114">Member</span></span> | <span data-ttu-id="579d9-115">Тип</span><span class="sxs-lookup"><span data-stu-id="579d9-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="579d9-116">accountType</span><span class="sxs-lookup"><span data-stu-id="579d9-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="579d9-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="579d9-117">Member</span></span> |
| [<span data-ttu-id="579d9-118">displayName</span><span class="sxs-lookup"><span data-stu-id="579d9-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="579d9-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="579d9-119">Member</span></span> |
| [<span data-ttu-id="579d9-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="579d9-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="579d9-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="579d9-121">Member</span></span> |
| [<span data-ttu-id="579d9-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="579d9-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="579d9-123">Член</span><span class="sxs-lookup"><span data-stu-id="579d9-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="579d9-124">Элементы</span><span class="sxs-lookup"><span data-stu-id="579d9-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="579d9-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="579d9-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="579d9-126">В настоящее время этот элемент поддерживается только в Outlook 2016 или более поздней версии для Mac (сборка 16.9.1212 или более поздняя версия).</span><span class="sxs-lookup"><span data-stu-id="579d9-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="579d9-127">Возвращает тип учетной записи пользователя, связанной с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="579d9-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="579d9-128">Возможные значения перечислены в таблице ниже.</span><span class="sxs-lookup"><span data-stu-id="579d9-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="579d9-129">Значение</span><span class="sxs-lookup"><span data-stu-id="579d9-129">Value</span></span> | <span data-ttu-id="579d9-130">Описание</span><span class="sxs-lookup"><span data-stu-id="579d9-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="579d9-131">Почтовый ящик размещен на локальном сервере Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="579d9-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="579d9-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="579d9-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="579d9-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="579d9-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="579d9-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="579d9-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="579d9-135">Тип</span><span class="sxs-lookup"><span data-stu-id="579d9-135">Type</span></span>

*   <span data-ttu-id="579d9-136">String</span><span class="sxs-lookup"><span data-stu-id="579d9-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="579d9-137">Требования</span><span class="sxs-lookup"><span data-stu-id="579d9-137">Requirements</span></span>

|<span data-ttu-id="579d9-138">Требование</span><span class="sxs-lookup"><span data-stu-id="579d9-138">Requirement</span></span>| <span data-ttu-id="579d9-139">Значение</span><span class="sxs-lookup"><span data-stu-id="579d9-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="579d9-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="579d9-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="579d9-141">1.6</span><span class="sxs-lookup"><span data-stu-id="579d9-141">1.6</span></span> |
|[<span data-ttu-id="579d9-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="579d9-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="579d9-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="579d9-143">ReadItem</span></span>|
|[<span data-ttu-id="579d9-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="579d9-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="579d9-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="579d9-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="579d9-146">Пример</span><span class="sxs-lookup"><span data-stu-id="579d9-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="579d9-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="579d9-147">displayName :String</span></span>

<span data-ttu-id="579d9-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="579d9-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="579d9-149">Тип</span><span class="sxs-lookup"><span data-stu-id="579d9-149">Type</span></span>

*   <span data-ttu-id="579d9-150">String</span><span class="sxs-lookup"><span data-stu-id="579d9-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="579d9-151">Требования</span><span class="sxs-lookup"><span data-stu-id="579d9-151">Requirements</span></span>

|<span data-ttu-id="579d9-152">Требование</span><span class="sxs-lookup"><span data-stu-id="579d9-152">Requirement</span></span>| <span data-ttu-id="579d9-153">Значение</span><span class="sxs-lookup"><span data-stu-id="579d9-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="579d9-154">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="579d9-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="579d9-155">1.0</span><span class="sxs-lookup"><span data-stu-id="579d9-155">1.0</span></span>|
|[<span data-ttu-id="579d9-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="579d9-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="579d9-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="579d9-157">ReadItem</span></span>|
|[<span data-ttu-id="579d9-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="579d9-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="579d9-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="579d9-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="579d9-160">Пример</span><span class="sxs-lookup"><span data-stu-id="579d9-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="579d9-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="579d9-161">emailAddress :String</span></span>

<span data-ttu-id="579d9-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="579d9-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="579d9-163">Тип</span><span class="sxs-lookup"><span data-stu-id="579d9-163">Type</span></span>

*   <span data-ttu-id="579d9-164">String</span><span class="sxs-lookup"><span data-stu-id="579d9-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="579d9-165">Требования</span><span class="sxs-lookup"><span data-stu-id="579d9-165">Requirements</span></span>

|<span data-ttu-id="579d9-166">Требование</span><span class="sxs-lookup"><span data-stu-id="579d9-166">Requirement</span></span>| <span data-ttu-id="579d9-167">Значение</span><span class="sxs-lookup"><span data-stu-id="579d9-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="579d9-168">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="579d9-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="579d9-169">1.0</span><span class="sxs-lookup"><span data-stu-id="579d9-169">1.0</span></span>|
|[<span data-ttu-id="579d9-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="579d9-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="579d9-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="579d9-171">ReadItem</span></span>|
|[<span data-ttu-id="579d9-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="579d9-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="579d9-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="579d9-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="579d9-174">Пример</span><span class="sxs-lookup"><span data-stu-id="579d9-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="579d9-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="579d9-175">timeZone :String</span></span>

<span data-ttu-id="579d9-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="579d9-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="579d9-177">Тип</span><span class="sxs-lookup"><span data-stu-id="579d9-177">Type</span></span>

*   <span data-ttu-id="579d9-178">String</span><span class="sxs-lookup"><span data-stu-id="579d9-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="579d9-179">Требования</span><span class="sxs-lookup"><span data-stu-id="579d9-179">Requirements</span></span>

|<span data-ttu-id="579d9-180">Требование</span><span class="sxs-lookup"><span data-stu-id="579d9-180">Requirement</span></span>| <span data-ttu-id="579d9-181">Значение</span><span class="sxs-lookup"><span data-stu-id="579d9-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="579d9-182">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="579d9-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="579d9-183">1.0</span><span class="sxs-lookup"><span data-stu-id="579d9-183">1.0</span></span>|
|[<span data-ttu-id="579d9-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="579d9-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="579d9-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="579d9-185">ReadItem</span></span>|
|[<span data-ttu-id="579d9-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="579d9-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="579d9-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="579d9-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="579d9-188">Пример</span><span class="sxs-lookup"><span data-stu-id="579d9-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
