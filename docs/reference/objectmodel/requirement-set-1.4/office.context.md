---
title: Office. Context — набор обязательных элементов 1,4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 01b03a425460acf5fd6f68214fd93d346920086e
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871141"
---
# <a name="context"></a><span data-ttu-id="07f46-102">context</span><span class="sxs-lookup"><span data-stu-id="07f46-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="07f46-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="07f46-103">[Office](Office.md).context</span></span>

<span data-ttu-id="07f46-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="07f46-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="07f46-106">Требования</span><span class="sxs-lookup"><span data-stu-id="07f46-106">Requirements</span></span>

|<span data-ttu-id="07f46-107">Требование</span><span class="sxs-lookup"><span data-stu-id="07f46-107">Requirement</span></span>| <span data-ttu-id="07f46-108">Значение</span><span class="sxs-lookup"><span data-stu-id="07f46-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="07f46-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="07f46-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07f46-110">1.0</span><span class="sxs-lookup"><span data-stu-id="07f46-110">1.0</span></span>|
|[<span data-ttu-id="07f46-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="07f46-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07f46-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="07f46-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="07f46-113">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="07f46-113">Namespaces</span></span>

<span data-ttu-id="07f46-114">[mailbox](office.context.mailbox.md). Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="07f46-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="07f46-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="07f46-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="07f46-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="07f46-116">displayLanguage :String</span></span>

<span data-ttu-id="07f46-117">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="07f46-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="07f46-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="07f46-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="07f46-119">Тип</span><span class="sxs-lookup"><span data-stu-id="07f46-119">Type</span></span>

*   <span data-ttu-id="07f46-120">String</span><span class="sxs-lookup"><span data-stu-id="07f46-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07f46-121">Требования</span><span class="sxs-lookup"><span data-stu-id="07f46-121">Requirements</span></span>

|<span data-ttu-id="07f46-122">Требование</span><span class="sxs-lookup"><span data-stu-id="07f46-122">Requirement</span></span>| <span data-ttu-id="07f46-123">Значение</span><span class="sxs-lookup"><span data-stu-id="07f46-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="07f46-124">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="07f46-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07f46-125">1.0</span><span class="sxs-lookup"><span data-stu-id="07f46-125">1.0</span></span>|
|[<span data-ttu-id="07f46-126">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="07f46-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07f46-127">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="07f46-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07f46-128">Пример</span><span class="sxs-lookup"><span data-stu-id="07f46-128">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="07f46-129">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="07f46-129">officeTheme :Object</span></span>

<span data-ttu-id="07f46-130">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="07f46-130">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="07f46-131">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="07f46-131">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="07f46-p102">Цвета тем Office позволяют согласовать цветовую схему надстройки с текущей темой Office, которую пользователь выбрал с помощью элементов **Файл > Учетная запись Office > Тема Office** и которая применяется во всех ведущих приложениях Office. Цвета тем Office можно использовать для всех надстроек почты и области задач.</span><span class="sxs-lookup"><span data-stu-id="07f46-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="07f46-134">Тип</span><span class="sxs-lookup"><span data-stu-id="07f46-134">Type</span></span>

*   <span data-ttu-id="07f46-135">Object</span><span class="sxs-lookup"><span data-stu-id="07f46-135">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="07f46-136">Свойства:</span><span class="sxs-lookup"><span data-stu-id="07f46-136">Properties:</span></span>

|<span data-ttu-id="07f46-137">Имя</span><span class="sxs-lookup"><span data-stu-id="07f46-137">Name</span></span>| <span data-ttu-id="07f46-138">Тип</span><span class="sxs-lookup"><span data-stu-id="07f46-138">Type</span></span>| <span data-ttu-id="07f46-139">Описание</span><span class="sxs-lookup"><span data-stu-id="07f46-139">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="07f46-140">String</span><span class="sxs-lookup"><span data-stu-id="07f46-140">String</span></span>|<span data-ttu-id="07f46-141">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="07f46-141">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="07f46-142">String</span><span class="sxs-lookup"><span data-stu-id="07f46-142">String</span></span>|<span data-ttu-id="07f46-143">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="07f46-143">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="07f46-144">String</span><span class="sxs-lookup"><span data-stu-id="07f46-144">String</span></span>|<span data-ttu-id="07f46-145">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="07f46-145">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="07f46-146">String</span><span class="sxs-lookup"><span data-stu-id="07f46-146">String</span></span>|<span data-ttu-id="07f46-147">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="07f46-147">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07f46-148">Требования</span><span class="sxs-lookup"><span data-stu-id="07f46-148">Requirements</span></span>

|<span data-ttu-id="07f46-149">Требование</span><span class="sxs-lookup"><span data-stu-id="07f46-149">Requirement</span></span>| <span data-ttu-id="07f46-150">Значение</span><span class="sxs-lookup"><span data-stu-id="07f46-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="07f46-151">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="07f46-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07f46-152">1.3</span><span class="sxs-lookup"><span data-stu-id="07f46-152">1.3</span></span>|
|[<span data-ttu-id="07f46-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="07f46-153">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07f46-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="07f46-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07f46-155">Пример</span><span class="sxs-lookup"><span data-stu-id="07f46-155">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook14officeroamingsettings"></a><span data-ttu-id="07f46-156">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="07f46-156">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span></span>

<span data-ttu-id="07f46-157">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="07f46-157">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="07f46-158">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="07f46-158">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="07f46-159">Тип</span><span class="sxs-lookup"><span data-stu-id="07f46-159">Type</span></span>

*   [<span data-ttu-id="07f46-160">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="07f46-160">RoamingSettings</span></span>](/javascript/api/outlook_1_4/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="07f46-161">Требования</span><span class="sxs-lookup"><span data-stu-id="07f46-161">Requirements</span></span>

|<span data-ttu-id="07f46-162">Требование</span><span class="sxs-lookup"><span data-stu-id="07f46-162">Requirement</span></span>| <span data-ttu-id="07f46-163">Значение</span><span class="sxs-lookup"><span data-stu-id="07f46-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="07f46-164">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="07f46-164">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07f46-165">1.0</span><span class="sxs-lookup"><span data-stu-id="07f46-165">1.0</span></span>|
|[<span data-ttu-id="07f46-166">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="07f46-166">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07f46-167">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="07f46-167">Restricted</span></span>|
|[<span data-ttu-id="07f46-168">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="07f46-168">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07f46-169">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="07f46-169">Compose or Read</span></span>|
