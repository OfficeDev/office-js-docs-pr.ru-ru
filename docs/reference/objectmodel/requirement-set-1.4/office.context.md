
# <a name="context"></a><span data-ttu-id="1dbf4-101">context</span><span class="sxs-lookup"><span data-stu-id="1dbf4-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="1dbf4-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="1dbf4-102">[Office](Office.md).context</span></span>

<span data-ttu-id="1dbf4-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="1dbf4-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dbf4-105">Требования</span><span class="sxs-lookup"><span data-stu-id="1dbf4-105">Requirements</span></span>

|<span data-ttu-id="1dbf4-106">Требование</span><span class="sxs-lookup"><span data-stu-id="1dbf4-106">Requirement</span></span>| <span data-ttu-id="1dbf4-107">Значение</span><span class="sxs-lookup"><span data-stu-id="1dbf4-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dbf4-108">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dbf4-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dbf4-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1dbf4-109">1.0</span></span>|
|[<span data-ttu-id="1dbf4-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dbf4-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dbf4-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1dbf4-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="1dbf4-112">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="1dbf4-112">Namespaces</span></span>

<span data-ttu-id="1dbf4-113">[mailbox](office.context.mailbox.md): предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-113">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="1dbf4-114">Члены</span><span class="sxs-lookup"><span data-stu-id="1dbf4-114">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="1dbf4-115">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="1dbf4-115">displayLanguage :String</span></span>

<span data-ttu-id="1dbf4-116">Получает определенный пользователем языковой стандарт в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-116">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="1dbf4-117">Значение `displayLanguage` отображает текущий параметр **Язык интерфейса**, заданный в разделе **Файл > Параметры > Язык** ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-117">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="1dbf4-118">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dbf4-118">Type:</span></span>

*   <span data-ttu-id="1dbf4-119">String</span><span class="sxs-lookup"><span data-stu-id="1dbf4-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1dbf4-120">Требования</span><span class="sxs-lookup"><span data-stu-id="1dbf4-120">Requirements</span></span>

|<span data-ttu-id="1dbf4-121">Требование</span><span class="sxs-lookup"><span data-stu-id="1dbf4-121">Requirement</span></span>| <span data-ttu-id="1dbf4-122">Значение</span><span class="sxs-lookup"><span data-stu-id="1dbf4-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dbf4-123">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dbf4-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dbf4-124">1.0</span><span class="sxs-lookup"><span data-stu-id="1dbf4-124">1.0</span></span>|
|[<span data-ttu-id="1dbf4-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dbf4-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dbf4-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1dbf4-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dbf4-127">Пример</span><span class="sxs-lookup"><span data-stu-id="1dbf4-127">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="1dbf4-128">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="1dbf4-128">officeTheme :Object</span></span>

<span data-ttu-id="1dbf4-129">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-129">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="1dbf4-130">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-130">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1dbf4-p102">Цвета тем Office позволяют согласовать цветовую схему надстройки с текущей темой Office, которую пользователь выбрал с помощью элементов **Файл > Учетная запись Office > Тема Office**, и которая применяется во всех ведущих приложениях Office. Цвета тем Office можно использовать для всех надстроек почты и области задач.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="1dbf4-133">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dbf4-133">Type:</span></span>

*   <span data-ttu-id="1dbf4-134">Object</span><span class="sxs-lookup"><span data-stu-id="1dbf4-134">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="1dbf4-135">Свойства:</span><span class="sxs-lookup"><span data-stu-id="1dbf4-135">Properties:</span></span>

|<span data-ttu-id="1dbf4-136">Name</span><span class="sxs-lookup"><span data-stu-id="1dbf4-136">Name</span></span>| <span data-ttu-id="1dbf4-137">Тип</span><span class="sxs-lookup"><span data-stu-id="1dbf4-137">Type</span></span>| <span data-ttu-id="1dbf4-138">Описание</span><span class="sxs-lookup"><span data-stu-id="1dbf4-138">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="1dbf4-139">String</span><span class="sxs-lookup"><span data-stu-id="1dbf4-139">String</span></span>|<span data-ttu-id="1dbf4-140">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-140">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="1dbf4-141">String</span><span class="sxs-lookup"><span data-stu-id="1dbf4-141">String</span></span>|<span data-ttu-id="1dbf4-142">Получает основной цвет текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-142">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="1dbf4-143">String</span><span class="sxs-lookup"><span data-stu-id="1dbf4-143">String</span></span>|<span data-ttu-id="1dbf4-144">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-144">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="1dbf4-145">String</span><span class="sxs-lookup"><span data-stu-id="1dbf4-145">String</span></span>|<span data-ttu-id="1dbf4-146">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-146">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1dbf4-147">Требования</span><span class="sxs-lookup"><span data-stu-id="1dbf4-147">Requirements</span></span>

|<span data-ttu-id="1dbf4-148">Требование</span><span class="sxs-lookup"><span data-stu-id="1dbf4-148">Requirement</span></span>| <span data-ttu-id="1dbf4-149">Значение</span><span class="sxs-lookup"><span data-stu-id="1dbf4-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dbf4-150">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dbf4-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dbf4-151">1.3</span><span class="sxs-lookup"><span data-stu-id="1dbf4-151">1.3</span></span>|
|[<span data-ttu-id="1dbf4-152">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dbf4-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dbf4-153">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1dbf4-153">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1dbf4-154">Пример</span><span class="sxs-lookup"><span data-stu-id="1dbf4-154">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook14officeroamingsettings"></a><span data-ttu-id="1dbf4-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="1dbf4-155">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_4/office.RoamingSettings)</span></span>

<span data-ttu-id="1dbf4-156">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-156">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="1dbf4-157">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="1dbf4-157">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="1dbf4-158">Тип:</span><span class="sxs-lookup"><span data-stu-id="1dbf4-158">Type:</span></span>

*   [<span data-ttu-id="1dbf4-159">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1dbf4-159">RoamingSettings</span></span>](/javascript/api/outlook_1_4/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="1dbf4-160">Требования</span><span class="sxs-lookup"><span data-stu-id="1dbf4-160">Requirements</span></span>

|<span data-ttu-id="1dbf4-161">Требование</span><span class="sxs-lookup"><span data-stu-id="1dbf4-161">Requirement</span></span>| <span data-ttu-id="1dbf4-162">Значение</span><span class="sxs-lookup"><span data-stu-id="1dbf4-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="1dbf4-163">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="1dbf4-163">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1dbf4-164">1.0</span><span class="sxs-lookup"><span data-stu-id="1dbf4-164">1.0</span></span>|
|[<span data-ttu-id="1dbf4-165">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1dbf4-165">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1dbf4-166">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="1dbf4-166">Restricted</span></span>|
|[<span data-ttu-id="1dbf4-167">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1dbf4-167">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1dbf4-168">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1dbf4-168">Compose or read</span></span>|