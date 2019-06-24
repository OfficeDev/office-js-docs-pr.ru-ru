---
title: Office. Context — набор обязательных элементов 1,3
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: ccc0a3addb5a440daf841452883019aa9f2b80c8
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127417"
---
# <a name="context"></a><span data-ttu-id="3bc2e-102">context</span><span class="sxs-lookup"><span data-stu-id="3bc2e-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="3bc2e-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="3bc2e-103">[Office](Office.md).context</span></span>

<span data-ttu-id="3bc2e-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="3bc2e-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3bc2e-106">Требования</span><span class="sxs-lookup"><span data-stu-id="3bc2e-106">Requirements</span></span>

|<span data-ttu-id="3bc2e-107">Требование</span><span class="sxs-lookup"><span data-stu-id="3bc2e-107">Requirement</span></span>| <span data-ttu-id="3bc2e-108">Значение</span><span class="sxs-lookup"><span data-stu-id="3bc2e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="3bc2e-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3bc2e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3bc2e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="3bc2e-110">1.0</span></span>|
|[<span data-ttu-id="3bc2e-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3bc2e-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3bc2e-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3bc2e-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="3bc2e-113">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="3bc2e-113">Namespaces</span></span>

<span data-ttu-id="3bc2e-114">[почтовый ящик](office.context.mailbox.md): предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="3bc2e-115">Members</span><span class="sxs-lookup"><span data-stu-id="3bc2e-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="3bc2e-116">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="3bc2e-116">displayLanguage: String</span></span>

<span data-ttu-id="3bc2e-117">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="3bc2e-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="3bc2e-119">Тип</span><span class="sxs-lookup"><span data-stu-id="3bc2e-119">Type</span></span>

*   <span data-ttu-id="3bc2e-120">String</span><span class="sxs-lookup"><span data-stu-id="3bc2e-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3bc2e-121">Требования</span><span class="sxs-lookup"><span data-stu-id="3bc2e-121">Requirements</span></span>

|<span data-ttu-id="3bc2e-122">Требование</span><span class="sxs-lookup"><span data-stu-id="3bc2e-122">Requirement</span></span>| <span data-ttu-id="3bc2e-123">Значение</span><span class="sxs-lookup"><span data-stu-id="3bc2e-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="3bc2e-124">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3bc2e-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3bc2e-125">1.0</span><span class="sxs-lookup"><span data-stu-id="3bc2e-125">1.0</span></span>|
|[<span data-ttu-id="3bc2e-126">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3bc2e-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3bc2e-127">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3bc2e-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3bc2e-128">Пример</span><span class="sxs-lookup"><span data-stu-id="3bc2e-128">Example</span></span>

```javascript
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

#### <a name="officetheme-object"></a><span data-ttu-id="3bc2e-129">officeTheme: объект</span><span class="sxs-lookup"><span data-stu-id="3bc2e-129">officeTheme: Object</span></span>

<span data-ttu-id="3bc2e-130">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-130">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="3bc2e-131">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-131">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3bc2e-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="3bc2e-134">Тип</span><span class="sxs-lookup"><span data-stu-id="3bc2e-134">Type</span></span>

*   <span data-ttu-id="3bc2e-135">Object</span><span class="sxs-lookup"><span data-stu-id="3bc2e-135">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="3bc2e-136">Свойства:</span><span class="sxs-lookup"><span data-stu-id="3bc2e-136">Properties:</span></span>

|<span data-ttu-id="3bc2e-137">Имя</span><span class="sxs-lookup"><span data-stu-id="3bc2e-137">Name</span></span>| <span data-ttu-id="3bc2e-138">Тип</span><span class="sxs-lookup"><span data-stu-id="3bc2e-138">Type</span></span>| <span data-ttu-id="3bc2e-139">Описание</span><span class="sxs-lookup"><span data-stu-id="3bc2e-139">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="3bc2e-140">String</span><span class="sxs-lookup"><span data-stu-id="3bc2e-140">String</span></span>|<span data-ttu-id="3bc2e-141">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-141">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="3bc2e-142">String</span><span class="sxs-lookup"><span data-stu-id="3bc2e-142">String</span></span>|<span data-ttu-id="3bc2e-143">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-143">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="3bc2e-144">String</span><span class="sxs-lookup"><span data-stu-id="3bc2e-144">String</span></span>|<span data-ttu-id="3bc2e-145">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-145">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="3bc2e-146">String</span><span class="sxs-lookup"><span data-stu-id="3bc2e-146">String</span></span>|<span data-ttu-id="3bc2e-147">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-147">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3bc2e-148">Требования</span><span class="sxs-lookup"><span data-stu-id="3bc2e-148">Requirements</span></span>

|<span data-ttu-id="3bc2e-149">Требование</span><span class="sxs-lookup"><span data-stu-id="3bc2e-149">Requirement</span></span>| <span data-ttu-id="3bc2e-150">Значение</span><span class="sxs-lookup"><span data-stu-id="3bc2e-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="3bc2e-151">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3bc2e-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3bc2e-152">1.3</span><span class="sxs-lookup"><span data-stu-id="3bc2e-152">1.3</span></span>|
|[<span data-ttu-id="3bc2e-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3bc2e-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3bc2e-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3bc2e-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3bc2e-155">Пример</span><span class="sxs-lookup"><span data-stu-id="3bc2e-155">Example</span></span>

```javascript
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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook13officeroamingsettings"></a><span data-ttu-id="3bc2e-156">roamingSettings: [roamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="3bc2e-156">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span></span>

<span data-ttu-id="3bc2e-157">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-157">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="3bc2e-158">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="3bc2e-158">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="3bc2e-159">Тип</span><span class="sxs-lookup"><span data-stu-id="3bc2e-159">Type</span></span>

*   [<span data-ttu-id="3bc2e-160">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="3bc2e-160">RoamingSettings</span></span>](/javascript/api/outlook_1_3/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="3bc2e-161">Требования</span><span class="sxs-lookup"><span data-stu-id="3bc2e-161">Requirements</span></span>

|<span data-ttu-id="3bc2e-162">Требование</span><span class="sxs-lookup"><span data-stu-id="3bc2e-162">Requirement</span></span>| <span data-ttu-id="3bc2e-163">Значение</span><span class="sxs-lookup"><span data-stu-id="3bc2e-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="3bc2e-164">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3bc2e-164">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3bc2e-165">1.0</span><span class="sxs-lookup"><span data-stu-id="3bc2e-165">1.0</span></span>|
|[<span data-ttu-id="3bc2e-166">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3bc2e-166">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3bc2e-167">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="3bc2e-167">Restricted</span></span>|
|[<span data-ttu-id="3bc2e-168">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3bc2e-168">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3bc2e-169">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3bc2e-169">Compose or Read</span></span>|
