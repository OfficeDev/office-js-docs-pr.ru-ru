---
title: Элемент DesktopFormFactor в файле манифеста
description: ''
ms.date: 05/08/2019
localization_priority: Normal
ms.openlocfilehash: b46536886d59692d03976083412a8b8d2e6ae859
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952392"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="97811-102">Элемент DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="97811-102">DesktopFormFactor element</span></span>

<span data-ttu-id="97811-103">Указывает параметры для надстройки классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="97811-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="97811-104">На настольных формах есть Office для Windows, Office для Mac и Office Online.</span><span class="sxs-lookup"><span data-stu-id="97811-104">The desktop form factor includes Office on Windows, Office for Mac, and Office Online.</span></span> <span data-ttu-id="97811-105">Он содержит все сведения о надстройке для классического форм-фактора, кроме узла **Resources**.</span><span class="sxs-lookup"><span data-stu-id="97811-105">It contains all the add-in information for the desktop form factor except for the  **Resources** node.</span></span>

<span data-ttu-id="97811-p102">В каждом определении DesktopFormFactor есть элемент **FunctionFile**, а также один или несколько элементов **ExtensionPoint**. Дополнительные сведения см. в статьях [Элемент FunctionFile](functionfile.md) и [Элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="97811-p102">Each DesktopFormFactor definition contains the  **FunctionFile** element and one or more **ExtensionPoint** elements. For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="97811-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="97811-108">Child elements</span></span>

| <span data-ttu-id="97811-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="97811-109">Element</span></span>                               | <span data-ttu-id="97811-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="97811-110">Required</span></span> | <span data-ttu-id="97811-111">Описание</span><span class="sxs-lookup"><span data-stu-id="97811-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="97811-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="97811-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="97811-113">Да</span><span class="sxs-lookup"><span data-stu-id="97811-113">Yes</span></span>      | <span data-ttu-id="97811-114">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="97811-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="97811-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="97811-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="97811-116">Да</span><span class="sxs-lookup"><span data-stu-id="97811-116">Yes</span></span>      | <span data-ttu-id="97811-117">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="97811-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="97811-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="97811-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="97811-119">Нет</span><span class="sxs-lookup"><span data-stu-id="97811-119">No</span></span>       | <span data-ttu-id="97811-120">Определяет выноску, которая отображается при установке надстройки в ведущих приложениях Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="97811-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="97811-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="97811-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="97811-122">Нет</span><span class="sxs-lookup"><span data-stu-id="97811-122">No</span></span> | <span data-ttu-id="97811-123">Определяет, доступна ли надстройка Outlook в сценариях делегирования, и имеет значение *false* по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="97811-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span><br><br><span data-ttu-id="97811-124">**Важно!** поскольку доступ представителя для надстроек Outlook в настоящее время находится в предварительной версии, надстройки, использующие `SupportSharedFolders` этот элемент, не могут быть опубликованы в AppSource или развернуты с помощью централизованного развертывания.</span><span class="sxs-lookup"><span data-stu-id="97811-124">**Important**: Because delegate access for Outlook add-ins is currently in preview, add-ins that use the `SupportSharedFolders` element cannot be published to AppSource or deployed via centralized deployment.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="97811-125">Пример DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="97811-125">DesktopFormFactor example</span></span>

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
