---
title: Элемент AllFormFactors в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: de7fcdce48e175d15ca6268f24082e37b2085b05
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433280"
---
# <a name="allformfactors-element"></a><span data-ttu-id="61f20-102">Элемент AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="61f20-102">AllFormFactors element</span></span>

<span data-ttu-id="61f20-103">Указывает параметры всех форм-факторов для надстройки.</span><span class="sxs-lookup"><span data-stu-id="61f20-103">Specifies the settings for an add-in for all form factors.</span></span> <span data-ttu-id="61f20-104">В настоящее время пользовательская функция — единственная, где применяется **AllFormFactors**.</span><span class="sxs-lookup"><span data-stu-id="61f20-104">Currently, the only feature using AllFormFactors is custom functions.</span></span> <span data-ttu-id="61f20-105">Элемент **AllFormFactors** является обязательным при использовании пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="61f20-105">AllFormFactors is a required element when using custom functions.</span></span>

## <a name="child-elements"></a><span data-ttu-id="61f20-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="61f20-106">Child elements</span></span>

|  <span data-ttu-id="61f20-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="61f20-107">Element</span></span> |  <span data-ttu-id="61f20-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="61f20-108">Required</span></span>  |  <span data-ttu-id="61f20-109">Описание</span><span class="sxs-lookup"><span data-stu-id="61f20-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="61f20-110">ExtensionPoint</span><span class="sxs-lookup"><span data-stu-id="61f20-110">ExtensionPoint</span></span>](extensionpoint.md) |  <span data-ttu-id="61f20-111">Да</span><span class="sxs-lookup"><span data-stu-id="61f20-111">Yes</span></span> |  <span data-ttu-id="61f20-112">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="61f20-112">Defines where an add-in exposes functionality.</span></span> |

## <a name="allformfactors-example"></a><span data-ttu-id="61f20-113">Пример использования AllFormFactors</span><span class="sxs-lookup"><span data-stu-id="61f20-113">AllFormFactors example</span></span>

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
