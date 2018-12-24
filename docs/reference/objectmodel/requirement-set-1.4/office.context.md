---
title: Office.context — набор обязательных элементов 1.4
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 9c868e26193b6b81284700af3df99bf165cd9d9e
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433588"
---
# <a name="context"></a><span data-ttu-id="ccfae-102">context</span><span class="sxs-lookup"><span data-stu-id="ccfae-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="ccfae-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="ccfae-103">[Office](Office.md).context</span></span>

<span data-ttu-id="ccfae-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="ccfae-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccfae-106">Требования</span><span class="sxs-lookup"><span data-stu-id="ccfae-106">Requirements</span></span>

|<span data-ttu-id="ccfae-107">Requirement</span><span class="sxs-lookup"><span data-stu-id="ccfae-107">Requirement</span></span>| <span data-ttu-id="ccfae-108">Значение</span><span class="sxs-lookup"><span data-stu-id="ccfae-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccfae-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ccfae-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccfae-110">1.0</span><span class="sxs-lookup"><span data-stu-id="ccfae-110">1.0</span></span>|
|[<span data-ttu-id="ccfae-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ccfae-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ccfae-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ccfae-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="ccfae-113">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="ccfae-113">Namespaces</span></span>

<span data-ttu-id="ccfae-114">[mailbox](office.context.mailbox.md). Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="ccfae-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="ccfae-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="ccfae-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="ccfae-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="ccfae-116">displayLanguage :String</span></span>

<span data-ttu-id="ccfae-117">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="ccfae-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="ccfae-118">Значение `displayLanguage` отображает текущий параметр **Язык интерфейса**, заданный в разделе **Файл > Параметры > Язык** ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="ccfae-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="ccfae-119">Тип:</span><span class="sxs-lookup"><span data-stu-id="ccfae-119">Type:</span></span>

*   <span data-ttu-id="ccfae-120">String</span><span class="sxs-lookup"><span data-stu-id="ccfae-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ccfae-121">Требования</span><span class="sxs-lookup"><span data-stu-id="ccfae-121">Requirements</span></span>

|<span data-ttu-id="ccfae-122">Requirement</span><span class="sxs-lookup"><span data-stu-id="ccfae-122">Requirement</span></span>| <span data-ttu-id="ccfae-123">Значение</span><span class="sxs-lookup"><span data-stu-id="ccfae-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccfae-124">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ccfae-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccfae-125">1.0</span><span class="sxs-lookup"><span data-stu-id="ccfae-125">1.0</span></span>|
|[<span data-ttu-id="ccfae-126">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ccfae-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ccfae-127">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ccfae-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccfae-128">Пример</span><span class="sxs-lookup"><span data-stu-id="ccfae-128">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="ccfae-129">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="ccfae-129">officeTheme :Object</span></span>

<span data-ttu-id="ccfae-130">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="ccfae-130">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="ccfae-131">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="ccfae-131">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ccfae-p102">Цвета тем Office позволяют согласовать цветовую схему надстройки с текущей темой Office, которую пользователь выбрал с помощью элементов **Файл > Учетная запись Office > Тема Office** и которая применяется во всех ведущих приложениях Office. Цвета тем Office можно использовать для всех надстроек почты и области задач.</span><span class="sxs-lookup"><span data-stu-id="ccfae-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ccfae-134">Тип:</span><span class="sxs-lookup"><span data-stu-id="ccfae-134">Type:</span></span>

*   <span data-ttu-id="ccfae-135">Object</span><span class="sxs-lookup"><span data-stu-id="ccfae-135">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="ccfae-136">Свойства:</span><span class="sxs-lookup"><span data-stu-id="ccfae-136">Properties:</span></span>

|<span data-ttu-id="ccfae-137">Имя</span><span class="sxs-lookup"><span data-stu-id="ccfae-137">Name</span></span>| <span data-ttu-id="ccfae-138">Тип</span><span class="sxs-lookup"><span data-stu-id="ccfae-138">Type</span></span>| <span data-ttu-id="ccfae-139">Описание</span><span class="sxs-lookup"><span data-stu-id="ccfae-139">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="ccfae-140">String</span><span class="sxs-lookup"><span data-stu-id="ccfae-140">String</span></span>|<span data-ttu-id="ccfae-141">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="ccfae-141">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="ccfae-142">String</span><span class="sxs-lookup"><span data-stu-id="ccfae-142">String</span></span>|<span data-ttu-id="ccfae-143">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="ccfae-143">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="ccfae-144">String</span><span class="sxs-lookup"><span data-stu-id="ccfae-144">String</span></span>|<span data-ttu-id="ccfae-145">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="ccfae-145">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="ccfae-146">String</span><span class="sxs-lookup"><span data-stu-id="ccfae-146">String</span></span>|<span data-ttu-id="ccfae-147">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="ccfae-147">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ccfae-148">Требования</span><span class="sxs-lookup"><span data-stu-id="ccfae-148">Requirements</span></span>

|<span data-ttu-id="ccfae-149">Requirement</span><span class="sxs-lookup"><span data-stu-id="ccfae-149">Requirement</span></span>| <span data-ttu-id="ccfae-150">Значение</span><span class="sxs-lookup"><span data-stu-id="ccfae-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccfae-151">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ccfae-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccfae-152">1.3</span><span class="sxs-lookup"><span data-stu-id="ccfae-152">1.3</span></span>|
|[<span data-ttu-id="ccfae-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ccfae-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ccfae-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ccfae-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ccfae-155">Пример</span><span class="sxs-lookup"><span data-stu-id="ccfae-155">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook14officeroamingsettings"></a><span data-ttu-id="ccfae-156">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="ccfae-156">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span></span>

<span data-ttu-id="ccfae-157">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="ccfae-157">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ccfae-158">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="ccfae-158">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ccfae-159">Тип:</span><span class="sxs-lookup"><span data-stu-id="ccfae-159">Type:</span></span>

*   [<span data-ttu-id="ccfae-160">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ccfae-160">RoamingSettings</span></span>](/javascript/api/outlook_1_4/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ccfae-161">Требования</span><span class="sxs-lookup"><span data-stu-id="ccfae-161">Requirements</span></span>

|<span data-ttu-id="ccfae-162">Requirement</span><span class="sxs-lookup"><span data-stu-id="ccfae-162">Requirement</span></span>| <span data-ttu-id="ccfae-163">Значение</span><span class="sxs-lookup"><span data-stu-id="ccfae-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="ccfae-164">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ccfae-164">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ccfae-165">1.0</span><span class="sxs-lookup"><span data-stu-id="ccfae-165">1.0</span></span>|
|[<span data-ttu-id="ccfae-166">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ccfae-166">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ccfae-167">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="ccfae-167">Restricted</span></span>|
|[<span data-ttu-id="ccfae-168">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ccfae-168">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ccfae-169">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ccfae-169">Compose or read</span></span>|