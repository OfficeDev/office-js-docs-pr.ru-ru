
# <a name="context"></a><span data-ttu-id="fc98f-101">context</span><span class="sxs-lookup"><span data-stu-id="fc98f-101">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="fc98f-102">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="fc98f-102">[Office](Office.md).context</span></span>

<span data-ttu-id="fc98f-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="fc98f-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="fc98f-105">Требования</span><span class="sxs-lookup"><span data-stu-id="fc98f-105">Requirements</span></span>

|<span data-ttu-id="fc98f-106">Требование</span><span class="sxs-lookup"><span data-stu-id="fc98f-106">Requirement</span></span>| <span data-ttu-id="fc98f-107">Значение</span><span class="sxs-lookup"><span data-stu-id="fc98f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc98f-108">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="fc98f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fc98f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="fc98f-109">1.0</span></span>|
|[<span data-ttu-id="fc98f-110">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fc98f-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fc98f-111">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fc98f-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="fc98f-112">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="fc98f-112">Members and methods</span></span>

| <span data-ttu-id="fc98f-113">Член</span><span class="sxs-lookup"><span data-stu-id="fc98f-113">Member</span></span> | <span data-ttu-id="fc98f-114">Тип</span><span class="sxs-lookup"><span data-stu-id="fc98f-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="fc98f-115">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="fc98f-115">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="fc98f-116">Член</span><span class="sxs-lookup"><span data-stu-id="fc98f-116">Member</span></span> |
| [<span data-ttu-id="fc98f-117">officeTheme</span><span class="sxs-lookup"><span data-stu-id="fc98f-117">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="fc98f-118">Член</span><span class="sxs-lookup"><span data-stu-id="fc98f-118">Member</span></span> |
| [<span data-ttu-id="fc98f-119">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="fc98f-119">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlook16officeroamingsettings) | <span data-ttu-id="fc98f-120">Член</span><span class="sxs-lookup"><span data-stu-id="fc98f-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="fc98f-121">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="fc98f-121">Namespaces</span></span>

<span data-ttu-id="fc98f-122">[mailbox](office.context.mailbox.md): предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="fc98f-122">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="fc98f-123">Члены</span><span class="sxs-lookup"><span data-stu-id="fc98f-123">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="fc98f-124">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="fc98f-124">displayLanguage :String</span></span>

<span data-ttu-id="fc98f-125">Получает определенный пользователем языковой стандарт в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="fc98f-125">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="fc98f-126">Значение `displayLanguage` отображает текущий параметр **Язык интерфейса**, заданный в разделе **Файл > Параметры > Язык** ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="fc98f-126">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="fc98f-127">Тип:</span><span class="sxs-lookup"><span data-stu-id="fc98f-127">Type:</span></span>

*   <span data-ttu-id="fc98f-128">String</span><span class="sxs-lookup"><span data-stu-id="fc98f-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fc98f-129">Требования</span><span class="sxs-lookup"><span data-stu-id="fc98f-129">Requirements</span></span>

|<span data-ttu-id="fc98f-130">Требование</span><span class="sxs-lookup"><span data-stu-id="fc98f-130">Requirement</span></span>| <span data-ttu-id="fc98f-131">Значение</span><span class="sxs-lookup"><span data-stu-id="fc98f-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc98f-132">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="fc98f-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fc98f-133">1.0</span><span class="sxs-lookup"><span data-stu-id="fc98f-133">1.0</span></span>|
|[<span data-ttu-id="fc98f-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fc98f-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fc98f-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fc98f-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fc98f-136">Пример</span><span class="sxs-lookup"><span data-stu-id="fc98f-136">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="fc98f-137">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="fc98f-137">officeTheme :Object</span></span>

<span data-ttu-id="fc98f-138">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="fc98f-138">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="fc98f-139">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="fc98f-139">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="fc98f-p102">Цвета тем Office позволяют согласовать цветовую схему надстройки с текущей темой Office, которую пользователь выбрал с помощью элементов **Файл > Учетная запись Office > Тема Office**, и которая применяется во всех ведущих приложениях Office. Цвета тем Office можно использовать для всех надстроек почты и области задач.</span><span class="sxs-lookup"><span data-stu-id="fc98f-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="fc98f-142">Тип:</span><span class="sxs-lookup"><span data-stu-id="fc98f-142">Type:</span></span>

*   <span data-ttu-id="fc98f-143">Объект</span><span class="sxs-lookup"><span data-stu-id="fc98f-143">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="fc98f-144">Свойства:</span><span class="sxs-lookup"><span data-stu-id="fc98f-144">Properties:</span></span>

|<span data-ttu-id="fc98f-145">Имя</span><span class="sxs-lookup"><span data-stu-id="fc98f-145">Name</span></span>| <span data-ttu-id="fc98f-146">Тип</span><span class="sxs-lookup"><span data-stu-id="fc98f-146">Type</span></span>| <span data-ttu-id="fc98f-147">Описание</span><span class="sxs-lookup"><span data-stu-id="fc98f-147">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="fc98f-148">String</span><span class="sxs-lookup"><span data-stu-id="fc98f-148">String</span></span>|<span data-ttu-id="fc98f-149">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="fc98f-149">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="fc98f-150">String</span><span class="sxs-lookup"><span data-stu-id="fc98f-150">String</span></span>|<span data-ttu-id="fc98f-151">Получает основной цвет текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="fc98f-151">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="fc98f-152">String</span><span class="sxs-lookup"><span data-stu-id="fc98f-152">String</span></span>|<span data-ttu-id="fc98f-153">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="fc98f-153">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="fc98f-154">String</span><span class="sxs-lookup"><span data-stu-id="fc98f-154">String</span></span>|<span data-ttu-id="fc98f-155">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="fc98f-155">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="fc98f-156">Требования</span><span class="sxs-lookup"><span data-stu-id="fc98f-156">Requirements</span></span>

|<span data-ttu-id="fc98f-157">Требование</span><span class="sxs-lookup"><span data-stu-id="fc98f-157">Requirement</span></span>| <span data-ttu-id="fc98f-158">Значение</span><span class="sxs-lookup"><span data-stu-id="fc98f-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc98f-159">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="fc98f-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fc98f-160">1.3</span><span class="sxs-lookup"><span data-stu-id="fc98f-160">1.3</span></span>|
|[<span data-ttu-id="fc98f-161">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fc98f-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fc98f-162">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fc98f-162">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fc98f-163">Пример</span><span class="sxs-lookup"><span data-stu-id="fc98f-163">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook16officeroamingsettings"></a><span data-ttu-id="fc98f-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_6/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="fc98f-164">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_6/office.RoamingSettings)</span></span>

<span data-ttu-id="fc98f-165">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="fc98f-165">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="fc98f-166">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="fc98f-166">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="fc98f-167">Тип:</span><span class="sxs-lookup"><span data-stu-id="fc98f-167">Type:</span></span>

*   [<span data-ttu-id="fc98f-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="fc98f-168">RoamingSettings</span></span>](/javascript/api/outlook_1_6/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="fc98f-169">Требования</span><span class="sxs-lookup"><span data-stu-id="fc98f-169">Requirements</span></span>

|<span data-ttu-id="fc98f-170">Требование</span><span class="sxs-lookup"><span data-stu-id="fc98f-170">Requirement</span></span>| <span data-ttu-id="fc98f-171">Значение</span><span class="sxs-lookup"><span data-stu-id="fc98f-171">Value</span></span>|
|---|---|
|[<span data-ttu-id="fc98f-172">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="fc98f-172">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fc98f-173">1.0</span><span class="sxs-lookup"><span data-stu-id="fc98f-173">1.0</span></span>|
|[<span data-ttu-id="fc98f-174">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="fc98f-174">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fc98f-175">Ограниченный доступ</span><span class="sxs-lookup"><span data-stu-id="fc98f-175">Restricted</span></span>|
|[<span data-ttu-id="fc98f-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="fc98f-176">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fc98f-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="fc98f-177">Compose or read</span></span>|