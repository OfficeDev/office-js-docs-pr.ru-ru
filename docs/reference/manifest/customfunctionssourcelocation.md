---
title: Элемент SourceLocation в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 878a8184984e31fdbcf46192a2f56507edaf4b37
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432410"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="cbb8c-102">Элемент SourceLocation</span><span class="sxs-lookup"><span data-stu-id="cbb8c-102">SourceLocation element</span></span>

<span data-ttu-id="cbb8c-103">Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="cbb8c-103">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="cbb8c-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="cbb8c-104">Attributes</span></span>

| <span data-ttu-id="cbb8c-105">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="cbb8c-105">**Attribute**</span></span> | <span data-ttu-id="cbb8c-106">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="cbb8c-106">**Required**</span></span> | <span data-ttu-id="cbb8c-107">**Описание**</span><span class="sxs-lookup"><span data-stu-id="cbb8c-107">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="cbb8c-108">resid</span><span class="sxs-lookup"><span data-stu-id="cbb8c-108">resid</span></span>         | <span data-ttu-id="cbb8c-109">Да</span><span class="sxs-lookup"><span data-stu-id="cbb8c-109">Yes</span></span>          | <span data-ttu-id="cbb8c-110">Имя ресурса URL-адреса, определенного в разделе &lt;Ресурсы&gt; в манифесте.</span><span class="sxs-lookup"><span data-stu-id="cbb8c-110">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="cbb8c-111">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="cbb8c-111">Child elements</span></span>

<span data-ttu-id="cbb8c-112">Нет</span><span class="sxs-lookup"><span data-stu-id="cbb8c-112">None</span></span>

## <a name="example"></a><span data-ttu-id="cbb8c-113">Пример</span><span class="sxs-lookup"><span data-stu-id="cbb8c-113">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```