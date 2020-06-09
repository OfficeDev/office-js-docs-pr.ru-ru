---
title: Элемент DesktopFormFactor в файле манифеста
description: Указывает параметры для надстройки классического форм-фактора.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 46de234f2d97a9e6c7645c17a0f0a61d0c3e1a80
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612285"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="1a088-103">Элемент DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="1a088-103">DesktopFormFactor element</span></span>

<span data-ttu-id="1a088-104">Указывает параметры для надстройки классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="1a088-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="1a088-105">Настольный конструктивный фактор включает Office в Интернете, Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="1a088-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="1a088-106">Он содержит все сведения о надстройках для настольных форм, за исключением узла **Resources** .</span><span class="sxs-lookup"><span data-stu-id="1a088-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="1a088-107">Каждое определение DesktopFormFactor содержит элемент **FunctionFile** и один или несколько элементов **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="1a088-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="1a088-108">Для получения дополнительных сведений см [элемент FunctionFile](functionfile.md) и [элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="1a088-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="1a088-109">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="1a088-109">Child elements</span></span>

| <span data-ttu-id="1a088-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="1a088-110">Element</span></span>                               | <span data-ttu-id="1a088-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="1a088-111">Required</span></span> | <span data-ttu-id="1a088-112">Описание</span><span class="sxs-lookup"><span data-stu-id="1a088-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="1a088-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="1a088-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="1a088-114">Да</span><span class="sxs-lookup"><span data-stu-id="1a088-114">Yes</span></span>      | <span data-ttu-id="1a088-115">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="1a088-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="1a088-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="1a088-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="1a088-117">Да</span><span class="sxs-lookup"><span data-stu-id="1a088-117">Yes</span></span>      | <span data-ttu-id="1a088-118">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="1a088-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="1a088-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="1a088-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="1a088-120">Нет</span><span class="sxs-lookup"><span data-stu-id="1a088-120">No</span></span>       | <span data-ttu-id="1a088-121">Определяет выноску, которая отображается при установке надстройки в ведущих приложениях Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="1a088-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="1a088-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="1a088-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="1a088-123">Нет</span><span class="sxs-lookup"><span data-stu-id="1a088-123">No</span></span> | <span data-ttu-id="1a088-124">Определяет, доступна ли надстройка Outlook в сценариях делегирования, и имеет значение *false* по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="1a088-124">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="1a088-125">Пример DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="1a088-125">DesktopFormFactor example</span></span>

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
