---
title: Элемент SourceLocation в файле манифеста
description: Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 56ebe122853c98a14c52d450bea31fecaefb15d3
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720689"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="53f35-103">Элемент SourceLocation</span><span class="sxs-lookup"><span data-stu-id="53f35-103">SourceLocation element</span></span>

<span data-ttu-id="53f35-104">Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="53f35-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="53f35-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="53f35-105">Attributes</span></span>

| <span data-ttu-id="53f35-106">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="53f35-106">**Attribute**</span></span> | <span data-ttu-id="53f35-107">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="53f35-107">**Required**</span></span> | <span data-ttu-id="53f35-108">**Описание**</span><span class="sxs-lookup"><span data-stu-id="53f35-108">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="53f35-109">resid</span><span class="sxs-lookup"><span data-stu-id="53f35-109">resid</span></span>         | <span data-ttu-id="53f35-110">Да</span><span class="sxs-lookup"><span data-stu-id="53f35-110">Yes</span></span>          | <span data-ttu-id="53f35-111">Имя ресурса URL-адреса, определенного в разделе &lt;Ресурсы&gt; в манифесте.</span><span class="sxs-lookup"><span data-stu-id="53f35-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="53f35-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="53f35-112">Child elements</span></span>

<span data-ttu-id="53f35-113">Нет</span><span class="sxs-lookup"><span data-stu-id="53f35-113">None</span></span>

## <a name="example"></a><span data-ttu-id="53f35-114">Пример</span><span class="sxs-lookup"><span data-stu-id="53f35-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
