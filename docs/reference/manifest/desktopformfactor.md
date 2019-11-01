---
title: Элемент DesktopFormFactor в файле манифеста
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: bada3cd4cff7973517aedb83235a224ef6c273eb
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/31/2019
ms.locfileid: "37901964"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="e8d4d-102">Элемент DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="e8d4d-102">DesktopFormFactor element</span></span>

<span data-ttu-id="e8d4d-103">Указывает параметры для надстройки классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="e8d4d-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="e8d4d-104">Настольный конструктивный фактор включает Office в Интернете, Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="e8d4d-104">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="e8d4d-105">Он содержит все сведения о надстройке для классического форм-фактора, кроме узла **Resources**.</span><span class="sxs-lookup"><span data-stu-id="e8d4d-105">It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="e8d4d-p102">В каждом определении DesktopFormFactor есть элемент **FunctionFile**, а также один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в статьях [Элемент FunctionFile](functionfile.md) и [Элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="e8d4d-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="e8d4d-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="e8d4d-108">Child elements</span></span>

| <span data-ttu-id="e8d4d-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="e8d4d-109">Element</span></span>                               | <span data-ttu-id="e8d4d-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e8d4d-110">Required</span></span> | <span data-ttu-id="e8d4d-111">Описание</span><span class="sxs-lookup"><span data-stu-id="e8d4d-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="e8d4d-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="e8d4d-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="e8d4d-113">Да</span><span class="sxs-lookup"><span data-stu-id="e8d4d-113">Yes</span></span>      | <span data-ttu-id="e8d4d-114">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="e8d4d-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="e8d4d-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="e8d4d-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="e8d4d-116">Да</span><span class="sxs-lookup"><span data-stu-id="e8d4d-116">Yes</span></span>      | <span data-ttu-id="e8d4d-117">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e8d4d-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="e8d4d-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="e8d4d-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="e8d4d-119">Нет</span><span class="sxs-lookup"><span data-stu-id="e8d4d-119">No</span></span>       | <span data-ttu-id="e8d4d-120">Определяет выноску, которая отображается при установке надстройки в ведущих приложениях Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="e8d4d-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="e8d4d-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="e8d4d-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="e8d4d-122">Нет</span><span class="sxs-lookup"><span data-stu-id="e8d4d-122">No</span></span> | <span data-ttu-id="e8d4d-123">Определяет, доступна ли надстройка Outlook в сценариях делегирования, и имеет значение *false* по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="e8d4d-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="e8d4d-124">Пример DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="e8d4d-124">DesktopFormFactor example</span></span>

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
        <!-- Information on this extension point. -->
      </ExtensionPoint>
      <!-- Possibly more ExtensionPoint elements. -->
    </DesktopFormFactor>
  </Host>
</Hosts>
...
```
