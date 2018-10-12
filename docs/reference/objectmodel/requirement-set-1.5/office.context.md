# <a name="context"></a><span data-ttu-id="721ea-101">context</span><span class="sxs-lookup"><span data-stu-id="721ea-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="721ea-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="721ea-102">[Office](Office.md).context</span></span>

<span data-ttu-id="721ea-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="721ea-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="721ea-105">Требования</span><span class="sxs-lookup"><span data-stu-id="721ea-105">Requirements</span></span>

|<span data-ttu-id="721ea-106">Требование</span><span class="sxs-lookup"><span data-stu-id="721ea-106">Requirement</span></span>| <span data-ttu-id="721ea-107">Значение</span><span class="sxs-lookup"><span data-stu-id="721ea-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="721ea-108">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="721ea-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="721ea-109">1.0</span><span class="sxs-lookup"><span data-stu-id="721ea-109">1.0</span></span>|
|[<span data-ttu-id="721ea-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="721ea-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="721ea-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="721ea-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="721ea-112">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="721ea-112">Members and methods</span></span>

| <span data-ttu-id="721ea-113">Член</span><span class="sxs-lookup"><span data-stu-id="721ea-113">Member</span></span> | <span data-ttu-id="721ea-114">Тип</span><span class="sxs-lookup"><span data-stu-id="721ea-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="721ea-115">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="721ea-115">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="721ea-116">Член</span><span class="sxs-lookup"><span data-stu-id="721ea-116">Member</span></span> |
| [<span data-ttu-id="721ea-117">officeTheme</span><span class="sxs-lookup"><span data-stu-id="721ea-117">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="721ea-118">Член</span><span class="sxs-lookup"><span data-stu-id="721ea-118">Member</span></span> |
| [<span data-ttu-id="721ea-119">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="721ea-119">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlook15officeroamingsettings) | <span data-ttu-id="721ea-120">Член</span><span class="sxs-lookup"><span data-stu-id="721ea-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="721ea-121">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="721ea-121">Namespaces</span></span>

<span data-ttu-id="721ea-122">[mailbox](office.context.mailbox.md): предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="721ea-122">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="721ea-123">Члены</span><span class="sxs-lookup"><span data-stu-id="721ea-123">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="721ea-124">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="721ea-124">displayLanguage :String</span></span>

<span data-ttu-id="721ea-125">Получает определенный пользователем языковой стандарт в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="721ea-125">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="721ea-126">Значение `displayLanguage` отображает текущий параметр **Язык интерфейса**, заданный в разделе **Файл > Параметры > Язык** ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="721ea-126">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="721ea-127">Тип:</span><span class="sxs-lookup"><span data-stu-id="721ea-127">Type:</span></span>

*   <span data-ttu-id="721ea-128">Строка</span><span class="sxs-lookup"><span data-stu-id="721ea-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="721ea-129">Требования</span><span class="sxs-lookup"><span data-stu-id="721ea-129">Requirements</span></span>

|<span data-ttu-id="721ea-130">Требование</span><span class="sxs-lookup"><span data-stu-id="721ea-130">Requirement</span></span>| <span data-ttu-id="721ea-131">Значение</span><span class="sxs-lookup"><span data-stu-id="721ea-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="721ea-132">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="721ea-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="721ea-133">1.0</span><span class="sxs-lookup"><span data-stu-id="721ea-133">1.0</span></span>|
|[<span data-ttu-id="721ea-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="721ea-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="721ea-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="721ea-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="721ea-136">Пример</span><span class="sxs-lookup"><span data-stu-id="721ea-136">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="721ea-137">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="721ea-137">officeTheme :Object</span></span>

<span data-ttu-id="721ea-138">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="721ea-138">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="721ea-139">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="721ea-139">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="721ea-p102">Цвета тем Office позволяют согласовать цветовую схему надстройки с текущей темой Office, которую пользователь выбрал с помощью элементов **Файл > Учетная запись Office > Тема Office**, и которая применяется во всех ведущих приложениях Office. Цвета тем Office можно использовать для всех надстроек почты и области задач.</span><span class="sxs-lookup"><span data-stu-id="721ea-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="721ea-142">Тип:</span><span class="sxs-lookup"><span data-stu-id="721ea-142">Type:</span></span>

*   <span data-ttu-id="721ea-143">Object</span><span class="sxs-lookup"><span data-stu-id="721ea-143">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="721ea-144">Свойства:</span><span class="sxs-lookup"><span data-stu-id="721ea-144">Properties:</span></span>

|<span data-ttu-id="721ea-145">Имя</span><span class="sxs-lookup"><span data-stu-id="721ea-145">Name</span></span>| <span data-ttu-id="721ea-146">Тип</span><span class="sxs-lookup"><span data-stu-id="721ea-146">Type</span></span>| <span data-ttu-id="721ea-147">Описание</span><span class="sxs-lookup"><span data-stu-id="721ea-147">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="721ea-148">Строка</span><span class="sxs-lookup"><span data-stu-id="721ea-148">String</span></span>|<span data-ttu-id="721ea-149">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="721ea-149">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="721ea-150">Строка</span><span class="sxs-lookup"><span data-stu-id="721ea-150">String</span></span>|<span data-ttu-id="721ea-151">Получает основной цвет текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="721ea-151">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="721ea-152">Строка</span><span class="sxs-lookup"><span data-stu-id="721ea-152">String</span></span>|<span data-ttu-id="721ea-153">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="721ea-153">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="721ea-154">Строка</span><span class="sxs-lookup"><span data-stu-id="721ea-154">String</span></span>|<span data-ttu-id="721ea-155">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="721ea-155">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="721ea-156">Требования</span><span class="sxs-lookup"><span data-stu-id="721ea-156">Requirements</span></span>

|<span data-ttu-id="721ea-157">Требование</span><span class="sxs-lookup"><span data-stu-id="721ea-157">Requirement</span></span>| <span data-ttu-id="721ea-158">Значение</span><span class="sxs-lookup"><span data-stu-id="721ea-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="721ea-159">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="721ea-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="721ea-160">1.3</span><span class="sxs-lookup"><span data-stu-id="721ea-160">1.3</span></span>|
|[<span data-ttu-id="721ea-161">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="721ea-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="721ea-162">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="721ea-162">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="721ea-163">Пример</span><span class="sxs-lookup"><span data-stu-id="721ea-163">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook15officeroamingsettings"></a><span data-ttu-id="721ea-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_5/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="721ea-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_5/office.RoamingSettings)</span></span>

<span data-ttu-id="721ea-165">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="721ea-165">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="721ea-166">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="721ea-166">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="721ea-167">Тип:</span><span class="sxs-lookup"><span data-stu-id="721ea-167">Type:</span></span>

*   [<span data-ttu-id="721ea-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="721ea-168">RoamingSettings</span></span>](/javascript/api/outlook_1_5/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="721ea-169">Требования</span><span class="sxs-lookup"><span data-stu-id="721ea-169">Requirements</span></span>

|<span data-ttu-id="721ea-170">Требование</span><span class="sxs-lookup"><span data-stu-id="721ea-170">Requirement</span></span>| <span data-ttu-id="721ea-171">Значение</span><span class="sxs-lookup"><span data-stu-id="721ea-171">Value</span></span>|
|---|---|
|[<span data-ttu-id="721ea-172">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="721ea-172">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="721ea-173">1.0</span><span class="sxs-lookup"><span data-stu-id="721ea-173">1.0</span></span>|
|[<span data-ttu-id="721ea-174">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="721ea-174">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="721ea-175">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="721ea-175">Restricted</span></span>|
|[<span data-ttu-id="721ea-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="721ea-176">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="721ea-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="721ea-177">Compose or read</span></span>|