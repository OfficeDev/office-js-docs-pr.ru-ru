# <a name="desktopformfactor-element"></a><span data-ttu-id="f59dd-101">Элемент DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="f59dd-101">DesktopFormFactor element</span></span>

<span data-ttu-id="f59dd-p101">Указывает параметры для надстройки классического форм-фактора. Классический форм-фактор включает Office для Windows, Office для Mac и Office Online. Он содержит все сведения о надстройке для классического форм-фактора, кроме узла **Resources**.</span><span class="sxs-lookup"><span data-stu-id="f59dd-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="f59dd-p102">В каждом определении DesktopFormFactor есть элемент **FunctionFile**, а также один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в статьях [Элемент FunctionFile](functionfile.md) и [Элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="f59dd-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f59dd-107">Элемент SupportsSharedFolders доступен только в наборе требований предварительной версии надстроек Outlook относительно Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="f59dd-107">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span>
> <span data-ttu-id="f59dd-108">Надстройки, использующие этот элемент, не разрешаются в магазине Office или для централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="f59dd-108">Add-ins that use this element aren't allowed in the Office Store or Centralized Deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="f59dd-109">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="f59dd-109">Child elements</span></span>

| <span data-ttu-id="f59dd-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="f59dd-110">Element</span></span>                               | <span data-ttu-id="f59dd-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="f59dd-111">Required</span></span> | <span data-ttu-id="f59dd-112">Описание</span><span class="sxs-lookup"><span data-stu-id="f59dd-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="f59dd-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="f59dd-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="f59dd-114">Да</span><span class="sxs-lookup"><span data-stu-id="f59dd-114">Yes</span></span>      | <span data-ttu-id="f59dd-115">Определяет, где предоставляется функциональность надстройки.</span><span class="sxs-lookup"><span data-stu-id="f59dd-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="f59dd-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="f59dd-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="f59dd-117">Да</span><span class="sxs-lookup"><span data-stu-id="f59dd-117">Yes</span></span>      | <span data-ttu-id="f59dd-118">URL-адрес файла, содержащего функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="f59dd-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="f59dd-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="f59dd-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="f59dd-120">Нет</span><span class="sxs-lookup"><span data-stu-id="f59dd-120">No</span></span>       | <span data-ttu-id="f59dd-121">Определяет выноску, которая отображается при установке надстройки в основных приложениях Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="f59dd-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| <span data-ttu-id="f59dd-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="f59dd-122">SupportsSharedFolders</span></span>                 | <span data-ttu-id="f59dd-123">Нет</span><span class="sxs-lookup"><span data-stu-id="f59dd-123">No</span></span>       | <span data-ttu-id="f59dd-124">Определяет, доступна ли надстройка Outlook в сценарии делегата, и имеет значение *false* по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="f59dd-124">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> <span data-ttu-id="f59dd-125">Набор требований для предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="f59dd-125">Outlook add-in API Preview requirement set</span></span>|

## <a name="desktopformfactor-example"></a><span data-ttu-id="f59dd-126">Пример DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="f59dd-126">DesktopFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="Presentation">
    <DesktopFormFactor>
      <FunctionFile resid="residDesktopFuncUrl" />
      <GetStarted>
        <!-- GetStarted callout -->
      </GetStarted>
      <ExtensionPoint xsi:type="PrimaryCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint>
      <!-- possibly more ExtensionPoint elements -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
