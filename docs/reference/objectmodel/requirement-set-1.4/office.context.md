---
title: Office. Context — набор обязательных элементов 1,4
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: ad1887f32568f30cb87e52dd1f9457be2022beb2
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127340"
---
# <a name="context"></a><span data-ttu-id="24924-102">context</span><span class="sxs-lookup"><span data-stu-id="24924-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="24924-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="24924-103">[Office](Office.md).context</span></span>

<span data-ttu-id="24924-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="24924-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="24924-106">Требования</span><span class="sxs-lookup"><span data-stu-id="24924-106">Requirements</span></span>

|<span data-ttu-id="24924-107">Требование</span><span class="sxs-lookup"><span data-stu-id="24924-107">Requirement</span></span>| <span data-ttu-id="24924-108">Значение</span><span class="sxs-lookup"><span data-stu-id="24924-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="24924-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="24924-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24924-110">1.0</span><span class="sxs-lookup"><span data-stu-id="24924-110">1.0</span></span>|
|[<span data-ttu-id="24924-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="24924-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24924-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="24924-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="24924-113">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="24924-113">Namespaces</span></span>

<span data-ttu-id="24924-114">[почтовый ящик](office.context.mailbox.md): предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="24924-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="24924-115">Members</span><span class="sxs-lookup"><span data-stu-id="24924-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="24924-116">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="24924-116">displayLanguage: String</span></span>

<span data-ttu-id="24924-117">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="24924-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="24924-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="24924-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="24924-119">Тип</span><span class="sxs-lookup"><span data-stu-id="24924-119">Type</span></span>

*   <span data-ttu-id="24924-120">String</span><span class="sxs-lookup"><span data-stu-id="24924-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="24924-121">Требования</span><span class="sxs-lookup"><span data-stu-id="24924-121">Requirements</span></span>

|<span data-ttu-id="24924-122">Требование</span><span class="sxs-lookup"><span data-stu-id="24924-122">Requirement</span></span>| <span data-ttu-id="24924-123">Значение</span><span class="sxs-lookup"><span data-stu-id="24924-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="24924-124">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="24924-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24924-125">1.0</span><span class="sxs-lookup"><span data-stu-id="24924-125">1.0</span></span>|
|[<span data-ttu-id="24924-126">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="24924-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24924-127">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="24924-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24924-128">Пример</span><span class="sxs-lookup"><span data-stu-id="24924-128">Example</span></span>

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

#### <a name="officetheme-object"></a><span data-ttu-id="24924-129">officeTheme: объект</span><span class="sxs-lookup"><span data-stu-id="24924-129">officeTheme: Object</span></span>

<span data-ttu-id="24924-130">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="24924-130">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="24924-131">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="24924-131">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="24924-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="24924-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="24924-134">Тип</span><span class="sxs-lookup"><span data-stu-id="24924-134">Type</span></span>

*   <span data-ttu-id="24924-135">Object</span><span class="sxs-lookup"><span data-stu-id="24924-135">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="24924-136">Свойства:</span><span class="sxs-lookup"><span data-stu-id="24924-136">Properties:</span></span>

|<span data-ttu-id="24924-137">Имя</span><span class="sxs-lookup"><span data-stu-id="24924-137">Name</span></span>| <span data-ttu-id="24924-138">Тип</span><span class="sxs-lookup"><span data-stu-id="24924-138">Type</span></span>| <span data-ttu-id="24924-139">Описание</span><span class="sxs-lookup"><span data-stu-id="24924-139">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="24924-140">String</span><span class="sxs-lookup"><span data-stu-id="24924-140">String</span></span>|<span data-ttu-id="24924-141">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="24924-141">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="24924-142">String</span><span class="sxs-lookup"><span data-stu-id="24924-142">String</span></span>|<span data-ttu-id="24924-143">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="24924-143">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="24924-144">String</span><span class="sxs-lookup"><span data-stu-id="24924-144">String</span></span>|<span data-ttu-id="24924-145">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="24924-145">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="24924-146">String</span><span class="sxs-lookup"><span data-stu-id="24924-146">String</span></span>|<span data-ttu-id="24924-147">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="24924-147">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="24924-148">Требования</span><span class="sxs-lookup"><span data-stu-id="24924-148">Requirements</span></span>

|<span data-ttu-id="24924-149">Требование</span><span class="sxs-lookup"><span data-stu-id="24924-149">Requirement</span></span>| <span data-ttu-id="24924-150">Значение</span><span class="sxs-lookup"><span data-stu-id="24924-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="24924-151">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="24924-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24924-152">1.3</span><span class="sxs-lookup"><span data-stu-id="24924-152">1.3</span></span>|
|[<span data-ttu-id="24924-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="24924-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24924-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="24924-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="24924-155">Пример</span><span class="sxs-lookup"><span data-stu-id="24924-155">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook14officeroamingsettings"></a><span data-ttu-id="24924-156">roamingSettings: [roamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="24924-156">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span></span>

<span data-ttu-id="24924-157">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="24924-157">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="24924-158">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="24924-158">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="24924-159">Тип</span><span class="sxs-lookup"><span data-stu-id="24924-159">Type</span></span>

*   [<span data-ttu-id="24924-160">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="24924-160">RoamingSettings</span></span>](/javascript/api/outlook_1_4/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="24924-161">Требования</span><span class="sxs-lookup"><span data-stu-id="24924-161">Requirements</span></span>

|<span data-ttu-id="24924-162">Требование</span><span class="sxs-lookup"><span data-stu-id="24924-162">Requirement</span></span>| <span data-ttu-id="24924-163">Значение</span><span class="sxs-lookup"><span data-stu-id="24924-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="24924-164">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="24924-164">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="24924-165">1.0</span><span class="sxs-lookup"><span data-stu-id="24924-165">1.0</span></span>|
|[<span data-ttu-id="24924-166">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="24924-166">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="24924-167">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="24924-167">Restricted</span></span>|
|[<span data-ttu-id="24924-168">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="24924-168">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="24924-169">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="24924-169">Compose or Read</span></span>|
