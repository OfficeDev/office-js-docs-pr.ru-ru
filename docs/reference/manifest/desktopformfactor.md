---
title: Элемент DesktopFormFactor в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dea632f7f8afa5d9b69f257798022e9e520e9394
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433742"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="ffb8c-102">Элемент DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="ffb8c-102">DesktopFormFactor element</span></span>

<span data-ttu-id="ffb8c-p101">Указывает параметры для надстройки классического форм-фактора. Классический форм-фактор включает Office для Windows, Office для Mac и Office Online. Он содержит все сведения о надстройке для классического форм-фактора, кроме узла **Resources**.</span><span class="sxs-lookup"><span data-stu-id="ffb8c-p101">Specifies the settings for an add-in for the desktop form factor. The desktop form factor includes Office for Windows, Office for Mac, and Office Online. It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="ffb8c-p102">В каждом определении DesktopFormFactor есть элемент **FunctionFile**, а также один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в статьях [Элемент FunctionFile](functionfile.md) и [Элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="ffb8c-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="ffb8c-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="ffb8c-108">Child elements</span></span>

| <span data-ttu-id="ffb8c-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="ffb8c-109">Element</span></span>                               | <span data-ttu-id="ffb8c-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ffb8c-110">Required</span></span> | <span data-ttu-id="ffb8c-111">Описание</span><span class="sxs-lookup"><span data-stu-id="ffb8c-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="ffb8c-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="ffb8c-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="ffb8c-113">Да</span><span class="sxs-lookup"><span data-stu-id="ffb8c-113">Yes</span></span>      | <span data-ttu-id="ffb8c-114">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="ffb8c-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="ffb8c-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="ffb8c-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="ffb8c-116">Да</span><span class="sxs-lookup"><span data-stu-id="ffb8c-116">Yes</span></span>      | <span data-ttu-id="ffb8c-117">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="ffb8c-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="ffb8c-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="ffb8c-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="ffb8c-119">Нет</span><span class="sxs-lookup"><span data-stu-id="ffb8c-119">No</span></span>       | <span data-ttu-id="ffb8c-120">Определяет выноску, которая отображается при установке надстройки в ведущих приложениях Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="ffb8c-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="ffb8c-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="ffb8c-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="ffb8c-122">Нет</span><span class="sxs-lookup"><span data-stu-id="ffb8c-122">No</span></span> | <span data-ttu-id="ffb8c-123">Определяет, доступна ли надстройка Outlook в сценариях делегирования, и имеет значение *false* по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ffb8c-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="ffb8c-124">**Важно!** Этот элемент доступен только в предварительной версии набора обязательных элементов надстроек Outlook для Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="ffb8c-124">**Important**: This element is only available in the Outlook add-ins Preview requirement set against Exchange Online.</span></span> <span data-ttu-id="ffb8c-125">Надстройки, использующие этот элемент, нельзя опубликовать в AppSource или развернуть с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="ffb8c-125">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="ffb8c-126">Пример DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="ffb8c-126">DesktopFormFactor example</span></span>

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
