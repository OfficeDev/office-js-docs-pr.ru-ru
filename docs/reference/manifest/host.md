# <a name="host-element"></a><span data-ttu-id="abf73-101">Элемент Host</span><span class="sxs-lookup"><span data-stu-id="abf73-101">Host element</span></span>

<span data-ttu-id="abf73-102">Определяет тип приложения Office, в котором следует активировать надстройку.</span><span class="sxs-lookup"><span data-stu-id="abf73-102">Specifies an individual Office application type where the add-in should activate.</span></span>

> [!IMPORTANT] 
> <span data-ttu-id="abf73-103">Синтаксис элемента **Host** зависит от того, задается ли элемент в [базовом манифесте](#basic-manifest) или в узле [VersionOverrides](#versionoverrides-node).</span><span class="sxs-lookup"><span data-stu-id="abf73-103">Important: The **Host** element syntax varies depending on whether the element is defined within the [basic manifest](#basic-manifest) or within the [VersionOverrides](#versionoverrides-node) node.</span></span> <span data-ttu-id="abf73-104">Однако функциональность в обоих случаях одинакова.</span><span class="sxs-lookup"><span data-stu-id="abf73-104">However, the functionality is the same.</span></span>  

## <a name="basic-manifest"></a><span data-ttu-id="abf73-105">Базовый манифест</span><span class="sxs-lookup"><span data-stu-id="abf73-105">Basic manifest</span></span>

<span data-ttu-id="abf73-106">Если основное приложение задается в базовом манифесте (в разделе [OfficeApp](officeapp.md)), то его тип определяется атрибутом `Name`.</span><span class="sxs-lookup"><span data-stu-id="abf73-106">When defined in the basic manifest (under [OfficeApp](officeapp.md)), the host type is determined by the `Name` attribute.</span></span>   

### <a name="attributes"></a><span data-ttu-id="abf73-107">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="abf73-107">Attributes</span></span>

| <span data-ttu-id="abf73-108">Атрибут</span><span class="sxs-lookup"><span data-stu-id="abf73-108">Attribute</span></span>     | <span data-ttu-id="abf73-109">Тип</span><span class="sxs-lookup"><span data-stu-id="abf73-109">Type</span></span>   | <span data-ttu-id="abf73-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="abf73-110">Required</span></span> | <span data-ttu-id="abf73-111">Описание</span><span class="sxs-lookup"><span data-stu-id="abf73-111">Description</span></span>                                      |
|:--------------|:-------|:---------|:-------------------------------------------------|
| [<span data-ttu-id="abf73-112">Имя</span><span class="sxs-lookup"><span data-stu-id="abf73-112">Name</span></span>](#name) | <span data-ttu-id="abf73-113">строка</span><span class="sxs-lookup"><span data-stu-id="abf73-113">string</span></span> | <span data-ttu-id="abf73-114">обязательный</span><span class="sxs-lookup"><span data-stu-id="abf73-114">required</span></span> | <span data-ttu-id="abf73-115">Имя типа основного приложения Office.</span><span class="sxs-lookup"><span data-stu-id="abf73-115">The name of the type of Office host application.</span></span> |

### <a name="name"></a><span data-ttu-id="abf73-116">Имя</span><span class="sxs-lookup"><span data-stu-id="abf73-116">Name</span></span>
<span data-ttu-id="abf73-p102">Определяет тип основного приложения, для которого предназначена эта надстройка. Поддерживаются такие значения:</span><span class="sxs-lookup"><span data-stu-id="abf73-p102">Specifies the Host type targeted by this add-in. The value must be one of the following:</span></span>

- <span data-ttu-id="abf73-119">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="abf73-119">`Document` (Word)</span></span>
- <span data-ttu-id="abf73-120">`Database` (Access)</span><span class="sxs-lookup"><span data-stu-id="abf73-120">`Database` (Access)</span></span>
- <span data-ttu-id="abf73-121">`Mailbox` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="abf73-121">`Mailbox` (Outlook)</span></span>
- <span data-ttu-id="abf73-122">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="abf73-122">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="abf73-123">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="abf73-123">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="abf73-124">`Project` (Project)</span><span class="sxs-lookup"><span data-stu-id="abf73-124">`Project` (Project)</span></span>
- <span data-ttu-id="abf73-125">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="abf73-125">`Workbook` (Excel)</span></span>

### <a name="example"></a><span data-ttu-id="abf73-126">Пример</span><span class="sxs-lookup"><span data-stu-id="abf73-126">Example</span></span>
```xml
<Hosts>
    <Host Name="Mailbox">
    </Host>
</Hosts>
```

## <a name="versionoverrides-node"></a><span data-ttu-id="abf73-127">Узел VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="abf73-127">VersionOverrides node</span></span>
<span data-ttu-id="abf73-128">Если основной элемент задается в узле [VersionOverrides](versionoverrides.md), его тип определяет атрибут `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="abf73-128">When defined in [VersionOverrides](versionoverrides.md), the host type is determined by the `xsi:type` attribute.</span></span> 

### <a name="attributes"></a><span data-ttu-id="abf73-129">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="abf73-129">Attributes</span></span>

|  <span data-ttu-id="abf73-130">Атрибут</span><span class="sxs-lookup"><span data-stu-id="abf73-130">Attribute</span></span>  |  <span data-ttu-id="abf73-131">Обязательный</span><span class="sxs-lookup"><span data-stu-id="abf73-131">Required</span></span>  |  <span data-ttu-id="abf73-132">Описание</span><span class="sxs-lookup"><span data-stu-id="abf73-132">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="abf73-133">xsi:type</span><span class="sxs-lookup"><span data-stu-id="abf73-133">xsi:type</span></span>](#xsitype)  |  <span data-ttu-id="abf73-134">Да</span><span class="sxs-lookup"><span data-stu-id="abf73-134">Yes</span></span>  | <span data-ttu-id="abf73-135">Описывает приложение Office, к которому применяются эти параметры.</span><span class="sxs-lookup"><span data-stu-id="abf73-135">Describes the Office host where these settings apply.</span></span>|

### <a name="child-elements"></a><span data-ttu-id="abf73-136">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="abf73-136">Child elements</span></span>

|  <span data-ttu-id="abf73-137">Элемент</span><span class="sxs-lookup"><span data-stu-id="abf73-137">Element</span></span> |  <span data-ttu-id="abf73-138">Обязательный</span><span class="sxs-lookup"><span data-stu-id="abf73-138">Required</span></span>  |  <span data-ttu-id="abf73-139">Описание</span><span class="sxs-lookup"><span data-stu-id="abf73-139">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="abf73-140">DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="abf73-140">DesktopFormFactor</span></span>](desktopformfactor.md)    |  <span data-ttu-id="abf73-141">Да</span><span class="sxs-lookup"><span data-stu-id="abf73-141">Yes</span></span>   |  <span data-ttu-id="abf73-142">Определяет параметры классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="abf73-142">Defines the settings for the desktop form factor.</span></span> |
|  [<span data-ttu-id="abf73-143">MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="abf73-143">MobileFormFactor</span></span>](mobileformfactor.md)    |  <span data-ttu-id="abf73-144">Нет</span><span class="sxs-lookup"><span data-stu-id="abf73-144">No</span></span>   |  <span data-ttu-id="abf73-p103">Определяет параметры форм-фактора мобильного устройства. **Примечание.** Этот элемент поддерживается только в Outlook для iOS.</span><span class="sxs-lookup"><span data-stu-id="abf73-p103">Defines the settings for the mobile form factor. **Note:** this element is only supported in Outlook for iOS.</span></span> |
|  [<span data-ttu-id="abf73-147">AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="abf73-147">AllFormFactors</span></span>](allformfactors.md)    |  <span data-ttu-id="abf73-148">Нет</span><span class="sxs-lookup"><span data-stu-id="abf73-148">No</span></span>   |  <span data-ttu-id="abf73-149">Определяет параметры всех форм-факторов.</span><span class="sxs-lookup"><span data-stu-id="abf73-149">Defines the settings for all form factors.</span></span> <span data-ttu-id="abf73-150">Используется только пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="abf73-150">Only used by custom functions in Excel.</span></span> |

### <a name="xsitype"></a><span data-ttu-id="abf73-151">xsi:type</span><span class="sxs-lookup"><span data-stu-id="abf73-151">xsi:type</span></span>

<span data-ttu-id="abf73-152">Указывает, к какому основному приложению Office (Word, Excel, PowerPoint, Outlook, OneNote) применяются содержащиеся параметры.</span><span class="sxs-lookup"><span data-stu-id="abf73-152">Controls which Office host (Word, Excel, PowerPoint, Outlook, OneNote) where the contained settings apply.</span></span> <span data-ttu-id="abf73-153">Допустимые значения:</span><span class="sxs-lookup"><span data-stu-id="abf73-153">The value must be one of the following:</span></span>

- <span data-ttu-id="abf73-154">`Document` (Word)</span><span class="sxs-lookup"><span data-stu-id="abf73-154">`Document` (Word)</span></span>
- <span data-ttu-id="abf73-155">`MailHost` (Outlook)</span><span class="sxs-lookup"><span data-stu-id="abf73-155">`MailHost` (Outlook)</span></span>    
- <span data-ttu-id="abf73-156">`Notebook` (OneNote)</span><span class="sxs-lookup"><span data-stu-id="abf73-156">`Notebook` (OneNote)</span></span>
- <span data-ttu-id="abf73-157">`Presentation` (PowerPoint)</span><span class="sxs-lookup"><span data-stu-id="abf73-157">`Presentation` (PowerPoint)</span></span>
- <span data-ttu-id="abf73-158">`Workbook` (Excel)</span><span class="sxs-lookup"><span data-stu-id="abf73-158">`Workbook` (Excel)</span></span>

## <a name="host-example"></a><span data-ttu-id="abf73-159">Пример основного приложения</span><span class="sxs-lookup"><span data-stu-id="abf73-159">Host example</span></span> 
```xml
<Hosts>
    <Host xsi:type="MailHost">
        <!-- Host Settings -->
    </Host>
</Hosts>
```
