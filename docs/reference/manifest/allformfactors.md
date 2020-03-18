---
title: Элемент AllFormFactors в файле манифеста
description: Указывает параметры всех форм-факторов для надстройки.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f1285f92b5eb89993e7fcfe79aab2325b86aca3d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720717"
---
# <a name="allformfactors-element"></a><span data-ttu-id="ef0f5-103">Элемент AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="ef0f5-103">AllFormFactors element</span></span>

<span data-ttu-id="ef0f5-104">Указывает параметры всех форм-факторов для надстройки.</span><span class="sxs-lookup"><span data-stu-id="ef0f5-104">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="ef0f5-105">В настоящее время пользовательская функция — единственная, где применяется **AllFormFactors**.</span><span class="sxs-lookup"><span data-stu-id="ef0f5-105">Currently, the only feature using **AllFormFactors** is custom functions.</span></span> <span data-ttu-id="ef0f5-106">Элемент **AllFormFactors** является обязательным при использовании пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="ef0f5-106">**AllFormFactors** is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="ef0f5-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="ef0f5-107">Child elements</span></span>

|  <span data-ttu-id="ef0f5-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="ef0f5-108">Element</span></span> |  <span data-ttu-id="ef0f5-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="ef0f5-109">Required</span></span>  |  <span data-ttu-id="ef0f5-110">Описание</span><span class="sxs-lookup"><span data-stu-id="ef0f5-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="ef0f5-111">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="ef0f5-111">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="ef0f5-112">Да</span><span class="sxs-lookup"><span data-stu-id="ef0f5-112">Yes</span></span> |  <span data-ttu-id="ef0f5-113">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="ef0f5-113">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="ef0f5-114">Пример использования AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="ef0f5-114">AllFormFactors example</span></span>

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
