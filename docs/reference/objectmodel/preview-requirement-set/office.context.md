---
title: Office.context — предварительная версия набора обязательных элементов
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: f06509e291325c635581d902d1f4f440bd255314
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696465"
---
# <a name="context"></a><span data-ttu-id="772dd-102">context</span><span class="sxs-lookup"><span data-stu-id="772dd-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="772dd-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="772dd-103">[Office](Office.md).context</span></span>

<span data-ttu-id="772dd-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="772dd-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="772dd-106">Требования</span><span class="sxs-lookup"><span data-stu-id="772dd-106">Requirements</span></span>

|<span data-ttu-id="772dd-107">Требование</span><span class="sxs-lookup"><span data-stu-id="772dd-107">Requirement</span></span>| <span data-ttu-id="772dd-108">Значение</span><span class="sxs-lookup"><span data-stu-id="772dd-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="772dd-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="772dd-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="772dd-110">1.0</span><span class="sxs-lookup"><span data-stu-id="772dd-110">1.0</span></span>|
|[<span data-ttu-id="772dd-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="772dd-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="772dd-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="772dd-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="772dd-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="772dd-113">Members and methods</span></span>

| <span data-ttu-id="772dd-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="772dd-114">Member</span></span> | <span data-ttu-id="772dd-115">Тип</span><span class="sxs-lookup"><span data-stu-id="772dd-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="772dd-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="772dd-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="772dd-117">Member</span><span class="sxs-lookup"><span data-stu-id="772dd-117">Member</span></span> |
| [<span data-ttu-id="772dd-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="772dd-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="772dd-119">Member</span><span class="sxs-lookup"><span data-stu-id="772dd-119">Member</span></span> |
| [<span data-ttu-id="772dd-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="772dd-120">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="772dd-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="772dd-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="772dd-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="772dd-122">Namespaces</span></span>

<span data-ttu-id="772dd-123">[почтовый ящик](office.context.mailbox.md): предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="772dd-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="772dd-124">Members</span><span class="sxs-lookup"><span data-stu-id="772dd-124">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="772dd-125">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="772dd-125">displayLanguage: String</span></span>

<span data-ttu-id="772dd-126">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="772dd-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="772dd-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="772dd-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="772dd-128">Тип</span><span class="sxs-lookup"><span data-stu-id="772dd-128">Type</span></span>

*   <span data-ttu-id="772dd-129">String</span><span class="sxs-lookup"><span data-stu-id="772dd-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="772dd-130">Требования</span><span class="sxs-lookup"><span data-stu-id="772dd-130">Requirements</span></span>

|<span data-ttu-id="772dd-131">Требование</span><span class="sxs-lookup"><span data-stu-id="772dd-131">Requirement</span></span>| <span data-ttu-id="772dd-132">Значение</span><span class="sxs-lookup"><span data-stu-id="772dd-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="772dd-133">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="772dd-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="772dd-134">1.0</span><span class="sxs-lookup"><span data-stu-id="772dd-134">1.0</span></span>|
|[<span data-ttu-id="772dd-135">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="772dd-135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="772dd-136">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="772dd-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="772dd-137">Пример</span><span class="sxs-lookup"><span data-stu-id="772dd-137">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="officetheme-object"></a><span data-ttu-id="772dd-138">officeTheme: объект</span><span class="sxs-lookup"><span data-stu-id="772dd-138">officeTheme: Object</span></span>

<span data-ttu-id="772dd-139">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="772dd-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="772dd-140">Этот элемент поддерживается только в Outlook для Windows.</span><span class="sxs-lookup"><span data-stu-id="772dd-140">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="772dd-141">Использование цветов тем Office позволяет координировать цветовую схему надстройки с текущей темой Office, выбранной пользователем с помощью **файла > учетной записи office > пользовательского интерфейса темы**Office, которая применяется ко всем ведущим приложениям Office.</span><span class="sxs-lookup"><span data-stu-id="772dd-141">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="772dd-142">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="772dd-142">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="772dd-143">Тип</span><span class="sxs-lookup"><span data-stu-id="772dd-143">Type</span></span>

*   <span data-ttu-id="772dd-144">Object</span><span class="sxs-lookup"><span data-stu-id="772dd-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="772dd-145">Свойства:</span><span class="sxs-lookup"><span data-stu-id="772dd-145">Properties:</span></span>

|<span data-ttu-id="772dd-146">Имя</span><span class="sxs-lookup"><span data-stu-id="772dd-146">Name</span></span>| <span data-ttu-id="772dd-147">Тип</span><span class="sxs-lookup"><span data-stu-id="772dd-147">Type</span></span>| <span data-ttu-id="772dd-148">Описание</span><span class="sxs-lookup"><span data-stu-id="772dd-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="772dd-149">String</span><span class="sxs-lookup"><span data-stu-id="772dd-149">String</span></span>|<span data-ttu-id="772dd-150">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="772dd-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="772dd-151">String.</span><span class="sxs-lookup"><span data-stu-id="772dd-151">String</span></span>|<span data-ttu-id="772dd-152">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="772dd-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="772dd-153">String.</span><span class="sxs-lookup"><span data-stu-id="772dd-153">String</span></span>|<span data-ttu-id="772dd-154">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="772dd-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="772dd-155">String</span><span class="sxs-lookup"><span data-stu-id="772dd-155">String</span></span>|<span data-ttu-id="772dd-156">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="772dd-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="772dd-157">Требования</span><span class="sxs-lookup"><span data-stu-id="772dd-157">Requirements</span></span>

|<span data-ttu-id="772dd-158">Требование</span><span class="sxs-lookup"><span data-stu-id="772dd-158">Requirement</span></span>| <span data-ttu-id="772dd-159">Значение</span><span class="sxs-lookup"><span data-stu-id="772dd-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="772dd-160">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="772dd-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="772dd-161">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="772dd-161">Preview</span></span>|
|[<span data-ttu-id="772dd-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="772dd-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="772dd-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="772dd-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="772dd-164">Пример</span><span class="sxs-lookup"><span data-stu-id="772dd-164">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="772dd-165">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="772dd-165">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span></span>

<span data-ttu-id="772dd-166">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="772dd-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="772dd-167">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="772dd-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="772dd-168">Тип</span><span class="sxs-lookup"><span data-stu-id="772dd-168">Type</span></span>

*   [<span data-ttu-id="772dd-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="772dd-169">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="772dd-170">Требования</span><span class="sxs-lookup"><span data-stu-id="772dd-170">Requirements</span></span>

|<span data-ttu-id="772dd-171">Требование</span><span class="sxs-lookup"><span data-stu-id="772dd-171">Requirement</span></span>| <span data-ttu-id="772dd-172">Значение</span><span class="sxs-lookup"><span data-stu-id="772dd-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="772dd-173">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="772dd-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="772dd-174">1.0</span><span class="sxs-lookup"><span data-stu-id="772dd-174">1.0</span></span>|
|[<span data-ttu-id="772dd-175">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="772dd-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="772dd-176">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="772dd-176">Restricted</span></span>|
|[<span data-ttu-id="772dd-177">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="772dd-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="772dd-178">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="772dd-178">Compose or Read</span></span>|
