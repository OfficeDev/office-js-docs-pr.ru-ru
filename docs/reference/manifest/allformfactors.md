---
title: Элемент AllFormFactors в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8059501f88f966b285398ac7cf243e6b0e4e44ea
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450739"
---
# <a name="allformfactors-element"></a><span data-ttu-id="d8411-102">Элемент AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="d8411-102">AllFormFactors element</span></span>

<span data-ttu-id="d8411-103">Указывает параметры всех форм-факторов для надстройки.</span><span class="sxs-lookup"><span data-stu-id="d8411-103">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="d8411-104">В настоящее время пользовательская функция — единственная, где применяется **AllFormFactors**.</span><span class="sxs-lookup"><span data-stu-id="d8411-104">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="d8411-105">Элемент **AllFormFactors** является обязательным при использовании пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="d8411-105">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="d8411-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="d8411-106">Child elements</span></span>

|  <span data-ttu-id="d8411-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="d8411-107">Element</span></span> |  <span data-ttu-id="d8411-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d8411-108">Required</span></span>  |  <span data-ttu-id="d8411-109">Описание</span><span class="sxs-lookup"><span data-stu-id="d8411-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="d8411-110">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="d8411-110">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="d8411-111">Да</span><span class="sxs-lookup"><span data-stu-id="d8411-111">Yes</span></span> |  <span data-ttu-id="d8411-112">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="d8411-112">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="d8411-113">Пример использования AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="d8411-113">AllFormFactors example</span></span>

```xml
<Hosts>
    <Host xsi:type="Workbook">
        <AllFormFactors>
            <ExtensionPoint xsi:type="CustomFunctions">
                    <!-- Information on this extension point -->
            </ExtensionPoint>
        </AllFormFactors>
    </Host>
</Hosts>
```
