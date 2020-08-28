---
title: Элемент DesktopFormFactor в файле манифеста
description: Указывает параметры для надстройки классического форм-фактора.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 18828e6b61a45ae2dc1528b3f7a54e664af09519
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292316"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="aaa05-103">Элемент DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="aaa05-103">DesktopFormFactor element</span></span>

<span data-ttu-id="aaa05-104">Указывает параметры для надстройки классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="aaa05-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="aaa05-105">Настольный конструктивный фактор включает Office в Интернете, Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="aaa05-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="aaa05-106">Он содержит все сведения о надстройках для настольных форм, за исключением узла **Resources** .</span><span class="sxs-lookup"><span data-stu-id="aaa05-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="aaa05-107">Каждое определение DesktopFormFactor содержит элемент **FunctionFile** и один или несколько элементов **ExtensionPoint** .</span><span class="sxs-lookup"><span data-stu-id="aaa05-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="aaa05-108">Для получения дополнительных сведений см [элемент FunctionFile](functionfile.md) и [элемент ExtensionPoint](extensionpoint.md).</span><span class="sxs-lookup"><span data-stu-id="aaa05-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="aaa05-109">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="aaa05-109">Child elements</span></span>

| <span data-ttu-id="aaa05-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="aaa05-110">Element</span></span>                               | <span data-ttu-id="aaa05-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="aaa05-111">Required</span></span> | <span data-ttu-id="aaa05-112">Описание</span><span class="sxs-lookup"><span data-stu-id="aaa05-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="aaa05-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="aaa05-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="aaa05-114">Да</span><span class="sxs-lookup"><span data-stu-id="aaa05-114">Yes</span></span>      | <span data-ttu-id="aaa05-115">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="aaa05-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="aaa05-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="aaa05-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="aaa05-117">Да</span><span class="sxs-lookup"><span data-stu-id="aaa05-117">Yes</span></span>      | <span data-ttu-id="aaa05-118">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="aaa05-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="aaa05-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="aaa05-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="aaa05-120">Нет</span><span class="sxs-lookup"><span data-stu-id="aaa05-120">No</span></span>       | <span data-ttu-id="aaa05-121">Определяет выноску, которая отображается при установке надстройки в Word, Excel или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="aaa05-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint.</span></span> |
| [<span data-ttu-id="aaa05-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="aaa05-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="aaa05-123">Нет</span><span class="sxs-lookup"><span data-stu-id="aaa05-123">No</span></span> | <span data-ttu-id="aaa05-124">Определяет, доступна ли надстройка Outlook в сценариях делегирования.</span><span class="sxs-lookup"><span data-stu-id="aaa05-124">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="aaa05-125">По умолчанию задано значение *false* .</span><span class="sxs-lookup"><span data-stu-id="aaa05-125">Set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="aaa05-126">Пример DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="aaa05-126">DesktopFormFactor example</span></span>

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
