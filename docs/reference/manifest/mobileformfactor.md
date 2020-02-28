---
title: Элемент MobileFormFactor в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 34106011cb855b6ac7c6d0fc21c16fd13e52b281
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324843"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="0b483-102">Элемент MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="0b483-102">MobileFormFactor element</span></span>

<span data-ttu-id="0b483-p101">Указывает параметры для надстройки в случае форм-фактора мобильного устройства. Содержит все сведения о надстройке для форм-фактора мобильного устройства, кроме узла **Resources**.</span><span class="sxs-lookup"><span data-stu-id="0b483-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="0b483-105">Каждое определение **MobileFormFactor** содержит элемент **FunctionFile** и один или несколько элементов **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="0b483-105">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="0b483-106">Для получения дополнительных сведений см [элемент FunctionFile](functionfile.md) и [элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="0b483-106">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="0b483-p103">Элемент **MobileFormFactor** определен в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение `VersionOverridesV1_1` атрибута `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="0b483-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="0b483-109">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="0b483-109">Child elements</span></span>

| <span data-ttu-id="0b483-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="0b483-110">Element</span></span>                               | <span data-ttu-id="0b483-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0b483-111">Required</span></span> | <span data-ttu-id="0b483-112">Описание</span><span class="sxs-lookup"><span data-stu-id="0b483-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="0b483-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="0b483-113">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="0b483-114">Да</span><span class="sxs-lookup"><span data-stu-id="0b483-114">Yes</span></span>      | <span data-ttu-id="0b483-115">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="0b483-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="0b483-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="0b483-116">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="0b483-117">Да</span><span class="sxs-lookup"><span data-stu-id="0b483-117">Yes</span></span>      | <span data-ttu-id="0b483-118">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0b483-118">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="0b483-119">Пример MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="0b483-119">MobileFormFactor example</span></span>

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
