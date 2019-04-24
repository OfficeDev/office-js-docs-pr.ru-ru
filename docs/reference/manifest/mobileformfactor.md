---
title: Элемент MobileFormFactor в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: aead8ea0b60130109c5537dc0017f3a9e3ef986f
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450571"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="c3fec-102">Элемент MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="c3fec-102">MobileFormFactor element</span></span>

<span data-ttu-id="c3fec-p101">Указывает параметры для надстройки в случае форм-фактора мобильного устройства. Содержит все сведения о надстройке для форм-фактора мобильного устройства, кроме узла **Resources**.</span><span class="sxs-lookup"><span data-stu-id="c3fec-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="c3fec-p102">Каждое определение **MobileFormFactor** содержит элемент **FunctionFile** и один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в разделах [Элемент FunctionFile](functionfile.md) и [Элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="c3fec-p102">Each **MobileFormFactor** definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="c3fec-p103">Элемент **MobileFormFactor** определен в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение `VersionOverridesV1_1` атрибута `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="c3fec-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c3fec-109">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c3fec-109">Child elements</span></span>

| <span data-ttu-id="c3fec-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="c3fec-110">Element</span></span>                               | <span data-ttu-id="c3fec-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c3fec-111">Required</span></span> | <span data-ttu-id="c3fec-112">Описание</span><span class="sxs-lookup"><span data-stu-id="c3fec-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="c3fec-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="c3fec-113">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="c3fec-114">Да</span><span class="sxs-lookup"><span data-stu-id="c3fec-114">Yes</span></span>      | <span data-ttu-id="c3fec-115">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="c3fec-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="c3fec-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="c3fec-116">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="c3fec-117">Да</span><span class="sxs-lookup"><span data-stu-id="c3fec-117">Yes</span></span>      | <span data-ttu-id="c3fec-118">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c3fec-118">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="c3fec-119">Пример MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="c3fec-119">MobileFormFactor example</span></span>

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
