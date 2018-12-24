---
title: Элемент Script в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 95e4cbadc35302b4f76108e0ff2a51d31ca89aac
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433140"
---
# <a name="script-element"></a><span data-ttu-id="5457e-102">Элемент Script</span><span class="sxs-lookup"><span data-stu-id="5457e-102">Script element</span></span>

<span data-ttu-id="5457e-103">Определяет параметры сценариев, используемых пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="5457e-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="5457e-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="5457e-104">Attributes</span></span>

<span data-ttu-id="5457e-105">Нет</span><span class="sxs-lookup"><span data-stu-id="5457e-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="5457e-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="5457e-106">Child elements</span></span>

|<span data-ttu-id="5457e-107">Элементы</span><span class="sxs-lookup"><span data-stu-id="5457e-107">Elements</span></span>  |  <span data-ttu-id="5457e-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="5457e-108">Required</span></span>  |  <span data-ttu-id="5457e-109">Описание</span><span class="sxs-lookup"><span data-stu-id="5457e-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="5457e-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="5457e-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="5457e-111">Да</span><span class="sxs-lookup"><span data-stu-id="5457e-111">Yes</span></span>  | <span data-ttu-id="5457e-112">Строка с идентификатором ресурса файла JavaScript, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="5457e-112">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="5457e-113">Пример</span><span class="sxs-lookup"><span data-stu-id="5457e-113">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
