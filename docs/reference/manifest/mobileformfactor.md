---
title: Элемент MobileFormFactor в файле манифеста
description: Элемент MobileFormFactor указывает параметры параметров формы мобильного устройства для надстройки.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 954fff5d1e701ce53a6ad82fa276c048ca6d6f3a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720591"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="c6d74-103">Элемент MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="c6d74-103">MobileFormFactor element</span></span>

<span data-ttu-id="c6d74-p101">Указывает параметры для надстройки в случае форм-фактора мобильного устройства. Содержит все сведения о надстройке для форм-фактора мобильного устройства, кроме узла **Resources**.</span><span class="sxs-lookup"><span data-stu-id="c6d74-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="c6d74-106">Каждое определение **MobileFormFactor** содержит элемент **FunctionFile** и один или несколько элементов **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="c6d74-106">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="c6d74-107">Для получения дополнительных сведений см [элемент FunctionFile](functionfile.md) и [элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="c6d74-107">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="c6d74-p103">Элемент **MobileFormFactor** определен в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение `VersionOverridesV1_1` атрибута `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="c6d74-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="c6d74-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c6d74-110">Child elements</span></span>

| <span data-ttu-id="c6d74-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="c6d74-111">Element</span></span>                               | <span data-ttu-id="c6d74-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c6d74-112">Required</span></span> | <span data-ttu-id="c6d74-113">Описание</span><span class="sxs-lookup"><span data-stu-id="c6d74-113">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="c6d74-114">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="c6d74-114">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="c6d74-115">Да</span><span class="sxs-lookup"><span data-stu-id="c6d74-115">Yes</span></span>      | <span data-ttu-id="c6d74-116">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="c6d74-116">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="c6d74-117">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="c6d74-117">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="c6d74-118">Да</span><span class="sxs-lookup"><span data-stu-id="c6d74-118">Yes</span></span>      | <span data-ttu-id="c6d74-119">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c6d74-119">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="c6d74-120">Пример MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="c6d74-120">MobileFormFactor example</span></span>

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
