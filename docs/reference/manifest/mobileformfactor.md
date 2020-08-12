---
title: Элемент MobileFormFactor в файле манифеста
description: Элемент MobileFormFactor указывает параметры параметров формы мобильного устройства для надстройки.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 5e52e66a2b97a32a19d42a4938dbeaed8f367478
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641475"
---
# <a name="mobileformfactor-element"></a><span data-ttu-id="9a17d-103">Элемент MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="9a17d-103">MobileFormFactor element</span></span>

<span data-ttu-id="9a17d-p101">Указывает параметры для надстройки в случае форм-фактора мобильного устройства. Содержит все сведения о надстройке для форм-фактора мобильного устройства, кроме узла **Resources**.</span><span class="sxs-lookup"><span data-stu-id="9a17d-p101">Specifies the settings for an add-in for the mobile form factor. It contains all the add-in information for the mobile form factor except for the **Resources** node.</span></span>

<span data-ttu-id="9a17d-106">Каждое определение **MobileFormFactor** содержит элемент **FunctionFile** и один или несколько элементов **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="9a17d-106">Each **MobileFormFactor** definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="9a17d-107">Для получения дополнительных сведений см [элемент FunctionFile](functionfile.md) и [элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="9a17d-107">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

<span data-ttu-id="9a17d-p103">Элемент **MobileFormFactor** определен в схеме 1.1 VersionOverrides. Содержащийся элемент [VersionOverrides](versionoverrides.md) должен иметь значение `VersionOverridesV1_1` атрибута `xsi:type`.</span><span class="sxs-lookup"><span data-stu-id="9a17d-p103">The **MobileFormFactor** element is defined in VersionOverrides schema 1.1. The containing [VersionOverrides](versionoverrides.md) element must have an `xsi:type` attribute value of `VersionOverridesV1_1`.</span></span>

## <a name="child-elements"></a><span data-ttu-id="9a17d-110">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="9a17d-110">Child elements</span></span>

| <span data-ttu-id="9a17d-111">Элемент</span><span class="sxs-lookup"><span data-stu-id="9a17d-111">Element</span></span>                             | <span data-ttu-id="9a17d-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="9a17d-112">Required</span></span> | <span data-ttu-id="9a17d-113">Описание</span><span class="sxs-lookup"><span data-stu-id="9a17d-113">Description</span></span>  |
|:------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="9a17d-114">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="9a17d-114">ExtensionPoint</span></span>](extensionpoint.md) | <span data-ttu-id="9a17d-115">Да</span><span class="sxs-lookup"><span data-stu-id="9a17d-115">Yes</span></span>      | <span data-ttu-id="9a17d-116">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="9a17d-116">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="9a17d-117">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="9a17d-117">FunctionFile</span></span>](functionfile.md)     | <span data-ttu-id="9a17d-118">Да</span><span class="sxs-lookup"><span data-stu-id="9a17d-118">Yes</span></span>      | <span data-ttu-id="9a17d-119">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="9a17d-119">A URL to a file that contains JavaScript functions.</span></span>|

## <a name="mobileformfactor-example"></a><span data-ttu-id="9a17d-120">Пример MobileFormFactor</span><span class="sxs-lookup"><span data-stu-id="9a17d-120">MobileFormFactor example</span></span>

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
