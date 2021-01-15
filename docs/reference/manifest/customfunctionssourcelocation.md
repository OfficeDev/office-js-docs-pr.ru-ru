---
title: Элемент SourceLocation для пользовательских функций в файле манифеста
description: Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 6001673f1954a4af2de66ff7611069c3fb402a13
ms.sourcegitcommit: 2f75a37de349251bc0e0fc402c5ae6dc5c3b8b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/06/2021
ms.locfileid: "49771384"
---
# <a name="sourcelocation-element-custom-functions"></a><span data-ttu-id="6f3e0-103">Элемент SourceLocation (пользовательские функции)</span><span class="sxs-lookup"><span data-stu-id="6f3e0-103">SourceLocation element (custom functions)</span></span>

<span data-ttu-id="6f3e0-104">Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="6f3e0-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="6f3e0-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="6f3e0-105">Attributes</span></span>

| <span data-ttu-id="6f3e0-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="6f3e0-106">Attribute</span></span> | <span data-ttu-id="6f3e0-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="6f3e0-107">Required</span></span> | <span data-ttu-id="6f3e0-108">Описание</span><span class="sxs-lookup"><span data-stu-id="6f3e0-108">Description</span></span>                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="6f3e0-109">resid</span><span class="sxs-lookup"><span data-stu-id="6f3e0-109">resid</span></span>     | <span data-ttu-id="6f3e0-110">Да</span><span class="sxs-lookup"><span data-stu-id="6f3e0-110">Yes</span></span>      | <span data-ttu-id="6f3e0-111">Имя ресурса URL-адреса, определенного в разделе &lt;Ресурсы&gt; в манифесте.</span><span class="sxs-lookup"><span data-stu-id="6f3e0-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> <span data-ttu-id="6f3e0-112">Может быть не более 32 символов.</span><span class="sxs-lookup"><span data-stu-id="6f3e0-112">Can be no more than 32 characters.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="6f3e0-113">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="6f3e0-113">Child elements</span></span>

<span data-ttu-id="6f3e0-114">Нет</span><span class="sxs-lookup"><span data-stu-id="6f3e0-114">None</span></span>

## <a name="example"></a><span data-ttu-id="6f3e0-115">Пример</span><span class="sxs-lookup"><span data-stu-id="6f3e0-115">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
