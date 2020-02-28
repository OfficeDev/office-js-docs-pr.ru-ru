---
title: Элемент DesktopFormFactor в файле манифеста
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 2fe97d99ff5bdc9f23a5760824e241ee4dfb800f
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325278"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="2f9e1-102">Элемент DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="2f9e1-102">DesktopFormFactor element</span></span>

<span data-ttu-id="2f9e1-103">Указывает параметры для надстройки классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="2f9e1-103">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="2f9e1-104">Настольный конструктивный фактор включает Office в Интернете, Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="2f9e1-104">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="2f9e1-105">Он содержит все сведения о надстройках для настольных форм, за исключением узла **Resources** .</span><span class="sxs-lookup"><span data-stu-id="2f9e1-105">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="2f9e1-106">Каждое определение DesktopFormFactor содержит элемент **FunctionFile** и один или несколько элементов **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="2f9e1-106">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="2f9e1-107">Для получения дополнительных сведений см [элемент FunctionFile](functionfile.md) и [элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="2f9e1-107">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="2f9e1-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="2f9e1-108">Child elements</span></span>

| <span data-ttu-id="2f9e1-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="2f9e1-109">Element</span></span>                               | <span data-ttu-id="2f9e1-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="2f9e1-110">Required</span></span> | <span data-ttu-id="2f9e1-111">Описание</span><span class="sxs-lookup"><span data-stu-id="2f9e1-111">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="2f9e1-112">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="2f9e1-112">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="2f9e1-113">Да</span><span class="sxs-lookup"><span data-stu-id="2f9e1-113">Yes</span></span>      | <span data-ttu-id="2f9e1-114">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="2f9e1-114">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="2f9e1-115">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="2f9e1-115">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="2f9e1-116">Да</span><span class="sxs-lookup"><span data-stu-id="2f9e1-116">Yes</span></span>      | <span data-ttu-id="2f9e1-117">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2f9e1-117">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="2f9e1-118">GetStarted</span><span class="sxs-lookup"><span data-stu-id="2f9e1-118">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="2f9e1-119">Нет</span><span class="sxs-lookup"><span data-stu-id="2f9e1-119">No</span></span>       | <span data-ttu-id="2f9e1-120">Определяет выноску, которая отображается при установке надстройки в ведущих приложениях Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="2f9e1-120">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint hosts.</span></span> |
| [<span data-ttu-id="2f9e1-121">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="2f9e1-121">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="2f9e1-122">Нет</span><span class="sxs-lookup"><span data-stu-id="2f9e1-122">No</span></span> | <span data-ttu-id="2f9e1-123">Определяет, доступна ли надстройка Outlook в сценариях делегирования, и имеет значение *false* по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="2f9e1-123">Defines whether the Outlook add-in is available in delegate scenarios and is set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="2f9e1-124">Пример DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="2f9e1-124">DesktopFormFactor example</span></span>

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
