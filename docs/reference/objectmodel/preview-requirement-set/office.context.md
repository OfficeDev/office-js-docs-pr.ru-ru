---
title: Office.context — предварительная версия набора обязательных элементов
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: afc74aba3ebacd2d8f54cdc94e02b9f9ba19dcbf
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068289"
---
# <a name="context"></a><span data-ttu-id="80149-102">context</span><span class="sxs-lookup"><span data-stu-id="80149-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="80149-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="80149-103">[Office](Office.md).context</span></span>

<span data-ttu-id="80149-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="80149-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="80149-106">Требования</span><span class="sxs-lookup"><span data-stu-id="80149-106">Requirements</span></span>

|<span data-ttu-id="80149-107">Требование</span><span class="sxs-lookup"><span data-stu-id="80149-107">Requirement</span></span>| <span data-ttu-id="80149-108">Значение</span><span class="sxs-lookup"><span data-stu-id="80149-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="80149-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="80149-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80149-110">1.0</span><span class="sxs-lookup"><span data-stu-id="80149-110">1.0</span></span>|
|[<span data-ttu-id="80149-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="80149-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80149-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="80149-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="80149-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="80149-113">Members and methods</span></span>

| <span data-ttu-id="80149-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="80149-114">Member</span></span> | <span data-ttu-id="80149-115">Тип</span><span class="sxs-lookup"><span data-stu-id="80149-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="80149-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="80149-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="80149-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="80149-117">Member</span></span> |
| [<span data-ttu-id="80149-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="80149-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="80149-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="80149-119">Member</span></span> |
| [<span data-ttu-id="80149-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="80149-120">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings) | <span data-ttu-id="80149-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="80149-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="80149-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="80149-122">Namespaces</span></span>

<span data-ttu-id="80149-123">[mailbox](office.context.mailbox.md). Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="80149-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="80149-124">Элементы</span><span class="sxs-lookup"><span data-stu-id="80149-124">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="80149-125">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="80149-125">displayLanguage :String</span></span>

<span data-ttu-id="80149-126">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="80149-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="80149-127">Значение `displayLanguage` отображает текущий параметр **Язык интерфейса**, заданный в разделе **Файл > Параметры > Язык** ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="80149-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="80149-128">Тип</span><span class="sxs-lookup"><span data-stu-id="80149-128">Type</span></span>

*   <span data-ttu-id="80149-129">String</span><span class="sxs-lookup"><span data-stu-id="80149-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="80149-130">Требования</span><span class="sxs-lookup"><span data-stu-id="80149-130">Requirements</span></span>

|<span data-ttu-id="80149-131">Требование</span><span class="sxs-lookup"><span data-stu-id="80149-131">Requirement</span></span>| <span data-ttu-id="80149-132">Значение</span><span class="sxs-lookup"><span data-stu-id="80149-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="80149-133">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="80149-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80149-134">1.0</span><span class="sxs-lookup"><span data-stu-id="80149-134">1.0</span></span>|
|[<span data-ttu-id="80149-135">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="80149-135">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80149-136">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="80149-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="80149-137">Пример</span><span class="sxs-lookup"><span data-stu-id="80149-137">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="80149-138">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="80149-138">officeTheme :Object</span></span>

<span data-ttu-id="80149-139">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="80149-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="80149-140">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="80149-140">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="80149-p102">Цвета тем Office позволяют согласовать цветовую схему надстройки с текущей темой Office, которую пользователь выбрал с помощью элементов **Файл > Учетная запись Office > Тема Office** и которая применяется во всех ведущих приложениях Office. Цвета тем Office можно использовать для всех надстроек почты и области задач.</span><span class="sxs-lookup"><span data-stu-id="80149-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="80149-143">Тип</span><span class="sxs-lookup"><span data-stu-id="80149-143">Type</span></span>

*   <span data-ttu-id="80149-144">Object</span><span class="sxs-lookup"><span data-stu-id="80149-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="80149-145">Свойства:</span><span class="sxs-lookup"><span data-stu-id="80149-145">Properties:</span></span>

|<span data-ttu-id="80149-146">Имя</span><span class="sxs-lookup"><span data-stu-id="80149-146">Name</span></span>| <span data-ttu-id="80149-147">Тип</span><span class="sxs-lookup"><span data-stu-id="80149-147">Type</span></span>| <span data-ttu-id="80149-148">Описание</span><span class="sxs-lookup"><span data-stu-id="80149-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="80149-149">String</span><span class="sxs-lookup"><span data-stu-id="80149-149">String</span></span>|<span data-ttu-id="80149-150">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="80149-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="80149-151">String</span><span class="sxs-lookup"><span data-stu-id="80149-151">String</span></span>|<span data-ttu-id="80149-152">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="80149-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="80149-153">String</span><span class="sxs-lookup"><span data-stu-id="80149-153">String</span></span>|<span data-ttu-id="80149-154">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="80149-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="80149-155">String</span><span class="sxs-lookup"><span data-stu-id="80149-155">String</span></span>|<span data-ttu-id="80149-156">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="80149-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="80149-157">Требования</span><span class="sxs-lookup"><span data-stu-id="80149-157">Requirements</span></span>

|<span data-ttu-id="80149-158">Требование</span><span class="sxs-lookup"><span data-stu-id="80149-158">Requirement</span></span>| <span data-ttu-id="80149-159">Значение</span><span class="sxs-lookup"><span data-stu-id="80149-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="80149-160">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="80149-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80149-161">1.3</span><span class="sxs-lookup"><span data-stu-id="80149-161">1.3</span></span>|
|[<span data-ttu-id="80149-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="80149-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80149-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="80149-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="80149-164">Пример</span><span class="sxs-lookup"><span data-stu-id="80149-164">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="80149-165">roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="80149-165">roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span></span>

<span data-ttu-id="80149-166">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="80149-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="80149-167">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="80149-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="80149-168">Тип</span><span class="sxs-lookup"><span data-stu-id="80149-168">Type</span></span>

*   [<span data-ttu-id="80149-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="80149-169">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="80149-170">Требования</span><span class="sxs-lookup"><span data-stu-id="80149-170">Requirements</span></span>

|<span data-ttu-id="80149-171">Требование</span><span class="sxs-lookup"><span data-stu-id="80149-171">Requirement</span></span>| <span data-ttu-id="80149-172">Значение</span><span class="sxs-lookup"><span data-stu-id="80149-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="80149-173">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="80149-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="80149-174">1.0</span><span class="sxs-lookup"><span data-stu-id="80149-174">1.0</span></span>|
|[<span data-ttu-id="80149-175">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="80149-175">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="80149-176">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="80149-176">Restricted</span></span>|
|[<span data-ttu-id="80149-177">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="80149-177">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="80149-178">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="80149-178">Compose or Read</span></span>|
