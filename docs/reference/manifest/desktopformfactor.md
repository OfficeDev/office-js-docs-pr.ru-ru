---
title: Элемент DesktopFormFactor в файле манифеста
description: Указывает параметры для надстройки классического форм-фактора.
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: 66673d83fd8608a1ec10492d7a944b0515de61c0
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007792"
---
# <a name="desktopformfactor-element"></a><span data-ttu-id="b81d9-103">Элемент DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="b81d9-103">DesktopFormFactor element</span></span>

<span data-ttu-id="b81d9-104">Указывает параметры для надстройки классического форм-фактора.</span><span class="sxs-lookup"><span data-stu-id="b81d9-104">Specifies the settings for an add-in for the desktop form factor.</span></span> <span data-ttu-id="b81d9-105">Форм-фактор рабочего стола включает Office в Интернете, Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="b81d9-105">The desktop form factor includes Office on the web, Windows, and Mac.</span></span> <span data-ttu-id="b81d9-106">Он содержит все сведения о надстройки для форм-фактора рабочего стола, за исключением **узла Resources.**</span><span class="sxs-lookup"><span data-stu-id="b81d9-106">It contains all the add-in information for the desktop form factor except for the **Resources** node.</span></span>

<span data-ttu-id="b81d9-107">Каждое определение DesktopFormFactor содержит элемент **FunctionFile** и один или несколько **элементов ExtensionPoint.**</span><span class="sxs-lookup"><span data-stu-id="b81d9-107">Each DesktopFormFactor definition contains the **FunctionFile** element and one or more **ExtensionPoint** elements.</span></span> <span data-ttu-id="b81d9-108">Дополнительные сведения см. в [элементе FunctionFile и](functionfile.md) [элементе ExtensionPoint.](extensionpoint.md)</span><span class="sxs-lookup"><span data-stu-id="b81d9-108">For more information, see [FunctionFile element](functionfile.md) and [ExtensionPoint element](extensionpoint.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="b81d9-109">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="b81d9-109">Child elements</span></span>

| <span data-ttu-id="b81d9-110">Элемент</span><span class="sxs-lookup"><span data-stu-id="b81d9-110">Element</span></span>                               | <span data-ttu-id="b81d9-111">Обязательный</span><span class="sxs-lookup"><span data-stu-id="b81d9-111">Required</span></span> | <span data-ttu-id="b81d9-112">Описание</span><span class="sxs-lookup"><span data-stu-id="b81d9-112">Description</span></span>  |
|:--------------------------------------|:--------:|:-------------|
| [<span data-ttu-id="b81d9-113">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="b81d9-113">ExtensionPoint</span></span>](extensionpoint.md)   | <span data-ttu-id="b81d9-114">Да</span><span class="sxs-lookup"><span data-stu-id="b81d9-114">Yes</span></span>      | <span data-ttu-id="b81d9-115">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="b81d9-115">Defines where an add-in exposes functionality.</span></span> |
| [<span data-ttu-id="b81d9-116">FunctionFile</span><span class="sxs-lookup"><span data-stu-id="b81d9-116">FunctionFile</span></span>](functionfile.md)       | <span data-ttu-id="b81d9-117">Да</span><span class="sxs-lookup"><span data-stu-id="b81d9-117">Yes</span></span>      | <span data-ttu-id="b81d9-118">URL-адрес файла, который содержит функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b81d9-118">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="b81d9-119">GetStarted</span><span class="sxs-lookup"><span data-stu-id="b81d9-119">GetStarted</span></span>](getstarted.md)           | <span data-ttu-id="b81d9-120">Нет</span><span class="sxs-lookup"><span data-stu-id="b81d9-120">No</span></span>       | <span data-ttu-id="b81d9-121">Определяет вызов, который появляется при установке надстройки в Word, Excel или PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="b81d9-121">Defines the callout that appears when installing the add-in in Word, Excel, or PowerPoint.</span></span> |
| [<span data-ttu-id="b81d9-122">SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="b81d9-122">SupportsSharedFolders</span></span>](supportssharedfolders.md) | <span data-ttu-id="b81d9-123">Нет</span><span class="sxs-lookup"><span data-stu-id="b81d9-123">No</span></span> | <span data-ttu-id="b81d9-124">Определяет, доступна ли надстройка Outlook в общих почтовых ящиках (в настоящее время в предварительном просмотре) и общих папках (т. е. в сценариях делегирования доступа).</span><span class="sxs-lookup"><span data-stu-id="b81d9-124">Defines whether the Outlook add-in is available in shared mailbox (now in preview) and shared folders (that is, delegate access) scenarios.</span></span> <span data-ttu-id="b81d9-125">Значение false *по* умолчанию.</span><span class="sxs-lookup"><span data-stu-id="b81d9-125">Set to *false* by default.</span></span> |

## <a name="desktopformfactor-example"></a><span data-ttu-id="b81d9-126">Пример DesktopFormFactor</span><span class="sxs-lookup"><span data-stu-id="b81d9-126">DesktopFormFactor example</span></span>

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
