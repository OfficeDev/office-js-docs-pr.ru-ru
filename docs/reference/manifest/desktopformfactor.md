# <a name="desktopformfactor-element"></a><span data-ttu-id="cebb9-101">Элемент DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="cebb9-101">DesktopFormFactor element</span></span>

<span data-ttu-id="cebb9-p101">Указывает параметры для надстройки классического форм-фактора. Классический форм-фактор включает Office для Windows, Office для Mac и Office Online. Он содержит все сведения о надстройке для классического форм-фактора, кроме узла **Resources**.</span><span class="sxs-lookup"><span data-stu-id="cebb9-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="cebb9-p102">В каждом определении DesktopFormFactor есть элемент **FunctionFile**, а также один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в статьях [Элемент FunctionFile](functionfile.md) и [Элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="cebb9-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="cebb9-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="cebb9-107">Child elements</span></span>

| <span data-ttu-id="cebb9-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="cebb9-108">Element</span></span>                               | <span data-ttu-id="cebb9-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="cebb9-109">Required</span></span> | <span data-ttu-id="cebb9-110">Описание</span><span class="sxs-lookup"><span data-stu-id="cebb9-110">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="cebb9-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="cebb9-111">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="cebb9-112">Да</span><span class="sxs-lookup"><span data-stu-id="cebb9-112">Yes</span></span>      | <span data-ttu-id="cebb9-113">Определяет, где предоставляется функциональность надстройки.</span><span class="sxs-lookup"><span data-stu-id="cebb9-113">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="cebb9-114">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="cebb9-114">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="cebb9-115">Да</span><span class="sxs-lookup"><span data-stu-id="cebb9-115">Yes</span></span>      | <span data-ttu-id="cebb9-116">URL-адрес файла, содержащего функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="cebb9-116">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="cebb9-117">GetStarted</span><span class="sxs-lookup"><span data-stu-id="cebb9-117">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="cebb9-118">Нет</span><span class="sxs-lookup"><span data-stu-id="cebb9-118">No</span></span>       | <span data-ttu-id="cebb9-119">Определяет выноску, которая отображается при установке надстройки в основных приложениях Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="cebb9-119">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="cebb9-120">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="cebb9-120">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="cebb9-121">Нет</span><span class="sxs-lookup"><span data-stu-id="cebb9-121">No</span></span> | <span data-ttu-id="cebb9-122">Определяет, доступна ли надстройка Outlook в сценарии делегата, и имеет значение *false* по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="cebb9-122">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="cebb9-123">**Важно**: этот элемент доступен только в наборе требований предварительной версии надстроек Outlook относительно Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="cebb9-123">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span> <span data-ttu-id="cebb9-124">Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="cebb9-124">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="cebb9-125">Пример DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="cebb9-125">DesktopFormFactor example</span></span>

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
