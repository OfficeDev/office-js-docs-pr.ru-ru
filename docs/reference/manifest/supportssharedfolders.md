# <a name="supportssharedfolders-element"></a><span data-ttu-id="892d3-101">Элемент SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="892d3-101">SupportsSharedFolders element</span></span>

<span data-ttu-id="892d3-102">Он определяет, является ли надстройка Outlook доступной в сценарии делегирования.</span><span class="sxs-lookup"><span data-stu-id="892d3-102">Defines whether the Outlook add-in is available in delegate scenarios and is set to false by default.</span></span> <span data-ttu-id="892d3-103">Элемент **SupportsSharedFolders** является дочерним элементом элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="892d3-103">The **ExtensionPoint** element is a child element of [AllFormFactors, DesktopFormFactor or MobileFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="892d3-104">Он имеет значение *false* по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="892d3-104">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="892d3-105">Этот элемент доступен только в [Наборе требований предварительного просмотра надстроек Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) по отношению к Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="892d3-105">The SupportsSharedFolders element is only available in the Outlook add-ins Preview Requirement Set against Exchange Online.</span></span> <span data-ttu-id="892d3-106">Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="892d3-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="892d3-107">Ниже приведен пример использования элемента **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="892d3-107">The following is an example of the **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <SupportsSharedFolders>true</SupportsSharedFolders>
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```
