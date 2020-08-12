---
title: Элемент SourceLocation для пользовательских функций в файле манифеста
description: Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.
ms.date: 08/07/2020
localization_priority: Normal
ms.openlocfilehash: 1c509987b0ce7948a63fa8ad51f7cf9c84144c5f
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641384"
---
# <a name="sourcelocation-element-custom-functions"></a><span data-ttu-id="d95b0-103">Элемент SourceLocation (пользовательские функции)</span><span class="sxs-lookup"><span data-stu-id="d95b0-103">SourceLocation element (custom functions)</span></span>

<span data-ttu-id="d95b0-104">Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="d95b0-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="d95b0-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d95b0-105">Attributes</span></span>

| <span data-ttu-id="d95b0-106">Атрибут</span><span class="sxs-lookup"><span data-stu-id="d95b0-106">Attribute</span></span> | <span data-ttu-id="d95b0-107">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d95b0-107">Required</span></span> | <span data-ttu-id="d95b0-108">Описание</span><span class="sxs-lookup"><span data-stu-id="d95b0-108">Description</span></span>                                                                          |
|-----------|----------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="d95b0-109">resid</span><span class="sxs-lookup"><span data-stu-id="d95b0-109">resid</span></span>     | <span data-ttu-id="d95b0-110">Да</span><span class="sxs-lookup"><span data-stu-id="d95b0-110">Yes</span></span>      | <span data-ttu-id="d95b0-111">Имя ресурса URL-адреса, определенного в разделе &lt;Ресурсы&gt; в манифесте.</span><span class="sxs-lookup"><span data-stu-id="d95b0-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="d95b0-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="d95b0-112">Child elements</span></span>

<span data-ttu-id="d95b0-113">Нет</span><span class="sxs-lookup"><span data-stu-id="d95b0-113">None</span></span>

## <a name="example"></a><span data-ttu-id="d95b0-114">Пример</span><span class="sxs-lookup"><span data-stu-id="d95b0-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
