---
title: Элемент AllFormFactors в файле манифеста
description: Указывает параметры всех форм-факторов для надстройки.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 9dac322312c1dfd60f6deb4296413e12b55a6a49
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608798"
---
# <a name="allformfactors-element"></a><span data-ttu-id="7520e-103">Элемент AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="7520e-103">AllFormFactors element</span></span>

<span data-ttu-id="7520e-104">Указывает параметры всех форм-факторов для надстройки.</span><span class="sxs-lookup"><span data-stu-id="7520e-104">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="7520e-105">В настоящее время пользовательская функция — единственная, где применяется **AllFormFactors**.</span><span class="sxs-lookup"><span data-stu-id="7520e-105">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="7520e-106">Элемент **AllFormFactors** является обязательным при использовании пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="7520e-106">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="7520e-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="7520e-107">Child elements</span></span>

|  <span data-ttu-id="7520e-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="7520e-108">Element</span></span> |  <span data-ttu-id="7520e-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="7520e-109">Required</span></span>  |  <span data-ttu-id="7520e-110">Описание</span><span class="sxs-lookup"><span data-stu-id="7520e-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="7520e-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="7520e-111">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="7520e-112">Да</span><span class="sxs-lookup"><span data-stu-id="7520e-112">Yes</span></span> |  <span data-ttu-id="7520e-113">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="7520e-113">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="7520e-114">Пример использования AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="7520e-114">AllFormFactors example</span></span>

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
