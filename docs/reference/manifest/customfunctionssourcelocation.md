---
title: Элемент SourceLocation в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: b2b78065fc8bde6fc827ddcb21e2bc700ed5bf49
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450690"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="c7b10-102">Элемент SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c7b10-102">SourceLocation element</span></span>

<span data-ttu-id="c7b10-103">Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="c7b10-103">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="c7b10-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c7b10-104">Attributes</span></span>

| <span data-ttu-id="c7b10-105">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="c7b10-105">**Attribute**</span></span> | <span data-ttu-id="c7b10-106">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="c7b10-106">**Required**</span></span> | <span data-ttu-id="c7b10-107">**Описание**</span><span class="sxs-lookup"><span data-stu-id="c7b10-107">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="c7b10-108">resid</span><span class="sxs-lookup"><span data-stu-id="c7b10-108">resid</span></span>         | <span data-ttu-id="c7b10-109">Да</span><span class="sxs-lookup"><span data-stu-id="c7b10-109">Yes</span></span>          | <span data-ttu-id="c7b10-110">Имя ресурса URL-адреса, определенного в разделе &lt;Ресурсы&gt; в манифесте.</span><span class="sxs-lookup"><span data-stu-id="c7b10-110">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="c7b10-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c7b10-111">Child elements</span></span>

<span data-ttu-id="c7b10-112">Нет</span><span class="sxs-lookup"><span data-stu-id="c7b10-112">None</span></span>

## <a name="example"></a><span data-ttu-id="c7b10-113">Пример</span><span class="sxs-lookup"><span data-stu-id="c7b10-113">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
