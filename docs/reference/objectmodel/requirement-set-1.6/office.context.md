---
title: Office. Context — набор обязательных элементов 1,6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 1b0e5e72ce6e9142a987b246a2034d9029654d45
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872013"
---
# <a name="context"></a><span data-ttu-id="836f7-102">context</span><span class="sxs-lookup"><span data-stu-id="836f7-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="836f7-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="836f7-103">[Office](Office.md).context</span></span>

<span data-ttu-id="836f7-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="836f7-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="836f7-106">Требования</span><span class="sxs-lookup"><span data-stu-id="836f7-106">Requirements</span></span>

|<span data-ttu-id="836f7-107">Требование</span><span class="sxs-lookup"><span data-stu-id="836f7-107">Requirement</span></span>| <span data-ttu-id="836f7-108">Значение</span><span class="sxs-lookup"><span data-stu-id="836f7-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="836f7-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="836f7-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="836f7-110">1.0</span><span class="sxs-lookup"><span data-stu-id="836f7-110">1.0</span></span>|
|[<span data-ttu-id="836f7-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="836f7-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="836f7-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="836f7-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="836f7-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="836f7-113">Members and methods</span></span>

| <span data-ttu-id="836f7-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="836f7-114">Member</span></span> | <span data-ttu-id="836f7-115">Тип</span><span class="sxs-lookup"><span data-stu-id="836f7-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="836f7-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="836f7-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="836f7-117">Member</span><span class="sxs-lookup"><span data-stu-id="836f7-117">Member</span></span> |
| [<span data-ttu-id="836f7-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="836f7-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="836f7-119">Member</span><span class="sxs-lookup"><span data-stu-id="836f7-119">Member</span></span> |
| [<span data-ttu-id="836f7-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="836f7-120">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="836f7-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="836f7-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="836f7-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="836f7-122">Namespaces</span></span>

<span data-ttu-id="836f7-123">[mailbox](office.context.mailbox.md). Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="836f7-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="836f7-124">Элементы</span><span class="sxs-lookup"><span data-stu-id="836f7-124">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="836f7-125">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="836f7-125">displayLanguage :String</span></span>

<span data-ttu-id="836f7-126">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="836f7-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="836f7-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="836f7-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="836f7-128">Тип</span><span class="sxs-lookup"><span data-stu-id="836f7-128">Type</span></span>

*   <span data-ttu-id="836f7-129">String</span><span class="sxs-lookup"><span data-stu-id="836f7-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="836f7-130">Требования</span><span class="sxs-lookup"><span data-stu-id="836f7-130">Requirements</span></span>

|<span data-ttu-id="836f7-131">Требование</span><span class="sxs-lookup"><span data-stu-id="836f7-131">Requirement</span></span>| <span data-ttu-id="836f7-132">Значение</span><span class="sxs-lookup"><span data-stu-id="836f7-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="836f7-133">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="836f7-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="836f7-134">1.0</span><span class="sxs-lookup"><span data-stu-id="836f7-134">1.0</span></span>|
|[<span data-ttu-id="836f7-135">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="836f7-135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="836f7-136">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="836f7-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="836f7-137">Пример</span><span class="sxs-lookup"><span data-stu-id="836f7-137">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="836f7-138">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="836f7-138">officeTheme :Object</span></span>

<span data-ttu-id="836f7-139">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="836f7-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="836f7-140">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="836f7-140">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="836f7-p102">Цвета тем Office позволяют согласовать цветовую схему надстройки с текущей темой Office, которую пользователь выбрал с помощью элементов **Файл > Учетная запись Office > Тема Office** и которая применяется во всех ведущих приложениях Office. Цвета тем Office можно использовать для всех надстроек почты и области задач.</span><span class="sxs-lookup"><span data-stu-id="836f7-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="836f7-143">Тип</span><span class="sxs-lookup"><span data-stu-id="836f7-143">Type</span></span>

*   <span data-ttu-id="836f7-144">Object</span><span class="sxs-lookup"><span data-stu-id="836f7-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="836f7-145">Свойства:</span><span class="sxs-lookup"><span data-stu-id="836f7-145">Properties:</span></span>

|<span data-ttu-id="836f7-146">Имя</span><span class="sxs-lookup"><span data-stu-id="836f7-146">Name</span></span>| <span data-ttu-id="836f7-147">Тип</span><span class="sxs-lookup"><span data-stu-id="836f7-147">Type</span></span>| <span data-ttu-id="836f7-148">Описание</span><span class="sxs-lookup"><span data-stu-id="836f7-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="836f7-149">String</span><span class="sxs-lookup"><span data-stu-id="836f7-149">String</span></span>|<span data-ttu-id="836f7-150">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="836f7-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="836f7-151">String</span><span class="sxs-lookup"><span data-stu-id="836f7-151">String</span></span>|<span data-ttu-id="836f7-152">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="836f7-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="836f7-153">String</span><span class="sxs-lookup"><span data-stu-id="836f7-153">String</span></span>|<span data-ttu-id="836f7-154">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="836f7-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="836f7-155">String</span><span class="sxs-lookup"><span data-stu-id="836f7-155">String</span></span>|<span data-ttu-id="836f7-156">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="836f7-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="836f7-157">Требования</span><span class="sxs-lookup"><span data-stu-id="836f7-157">Requirements</span></span>

|<span data-ttu-id="836f7-158">Требование</span><span class="sxs-lookup"><span data-stu-id="836f7-158">Requirement</span></span>| <span data-ttu-id="836f7-159">Значение</span><span class="sxs-lookup"><span data-stu-id="836f7-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="836f7-160">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="836f7-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="836f7-161">1.3</span><span class="sxs-lookup"><span data-stu-id="836f7-161">1.3</span></span>|
|[<span data-ttu-id="836f7-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="836f7-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="836f7-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="836f7-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="836f7-164">Пример</span><span class="sxs-lookup"><span data-stu-id="836f7-164">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook16officeroamingsettings"></a><span data-ttu-id="836f7-165">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_6/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="836f7-165">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_6/office.RoamingSettings)</span></span>

<span data-ttu-id="836f7-166">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="836f7-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="836f7-167">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="836f7-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="836f7-168">Тип</span><span class="sxs-lookup"><span data-stu-id="836f7-168">Type</span></span>

*   [<span data-ttu-id="836f7-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="836f7-169">RoamingSettings</span></span>](/javascript/api/outlook_1_6/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="836f7-170">Требования</span><span class="sxs-lookup"><span data-stu-id="836f7-170">Requirements</span></span>

|<span data-ttu-id="836f7-171">Требование</span><span class="sxs-lookup"><span data-stu-id="836f7-171">Requirement</span></span>| <span data-ttu-id="836f7-172">Значение</span><span class="sxs-lookup"><span data-stu-id="836f7-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="836f7-173">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="836f7-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="836f7-174">1.0</span><span class="sxs-lookup"><span data-stu-id="836f7-174">1.0</span></span>|
|[<span data-ttu-id="836f7-175">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="836f7-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="836f7-176">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="836f7-176">Restricted</span></span>|
|[<span data-ttu-id="836f7-177">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="836f7-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="836f7-178">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="836f7-178">Compose or Read</span></span>|
