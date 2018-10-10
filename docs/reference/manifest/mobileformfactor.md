# <a name="mobileformfactor-element"></a><span data-ttu-id="edf7a-101">Элемент MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="edf7a-101">MobileFormFactor element</span></span>

<span data-ttu-id="edf7a-p101">Указывает параметры для надстройки в случае форм-фактора мобильного устройства. Содержит все сведения о надстройке для форм-фактора мобильного устройства, кроме узла **Resources**.</span><span class="sxs-lookup"><span data-stu-id="edf7a-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="edf7a-p102">Каждое определение **MobileFormFactor** содержит элемент **FunctionFile** и один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в разделах [Элемент FunctionFile](functionfile.md) и [Элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="edf7a-p102">Each **MobileFormFactor** definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="edf7a-p103">Элемент **MobileFormFactor** определен в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение `VersionOverridesV1_1` атрибута `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="edf7a-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="edf7a-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="edf7a-108">Child elements</span></span>

| <span data-ttu-id="edf7a-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="edf7a-109">Element</span></span>                               | <span data-ttu-id="edf7a-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="edf7a-110">Required</span></span> | <span data-ttu-id="edf7a-111">Описание</span><span class="sxs-lookup"><span data-stu-id="edf7a-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="edf7a-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="edf7a-112">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="edf7a-113">Да</span><span class="sxs-lookup"><span data-stu-id="edf7a-113">Yes</span></span>      | <span data-ttu-id="edf7a-114">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="edf7a-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="edf7a-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="edf7a-115">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="edf7a-116">Да</span><span class="sxs-lookup"><span data-stu-id="edf7a-116">Yes</span></span>      | <span data-ttu-id="edf7a-117">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="edf7a-117">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="edf7a-118">Пример MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="edf7a-118">MobileFormFactor example</span></span>

```xml
...
<Hosts>
  <Host xsi:type="MailHost">
    ...
    <MobileFormFactor>
      <FunctionFile resid="residUILessFunctionFileUrl" />
      <ExtensionPoint xsi:type="MobileMessageReadCommandSurface">
        <!-- information on this extension point -->
      </ExtensionPoint> 
      <!-- possibly more ExtensionPoint elements -->
    </MobileFormFactor>
  </Host>
</Hosts>
...
```
