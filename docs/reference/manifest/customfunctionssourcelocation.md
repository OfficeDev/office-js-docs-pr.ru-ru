---
title: Элемент SourceLocation в файле манифеста
description: Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 88ae0558577167074a870170833617c4f60730f1
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612314"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="c6221-103">Элемент SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c6221-103">SourceLocation element</span></span>

<span data-ttu-id="c6221-104">Определяет расположение ресурса, который необходим для элементов Script или Page, используемых пользовательскими функциями в Excel.</span><span class="sxs-lookup"><span data-stu-id="c6221-104">Defines the location of a resource needed by the Script or Page elements used by custom functions in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="c6221-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c6221-105">Attributes</span></span>

| <span data-ttu-id="c6221-106">**Атрибут**</span><span class="sxs-lookup"><span data-stu-id="c6221-106">**Attribute**</span></span> | <span data-ttu-id="c6221-107">**Обязательный**</span><span class="sxs-lookup"><span data-stu-id="c6221-107">**Required**</span></span> | <span data-ttu-id="c6221-108">**Описание**</span><span class="sxs-lookup"><span data-stu-id="c6221-108">**Description**</span></span>                                                                      |
|---------------|--------------|--------------------------------------------------------------------------------------|
| <span data-ttu-id="c6221-109">resid</span><span class="sxs-lookup"><span data-stu-id="c6221-109">resid</span></span>         | <span data-ttu-id="c6221-110">Да</span><span class="sxs-lookup"><span data-stu-id="c6221-110">Yes</span></span>          | <span data-ttu-id="c6221-111">Имя ресурса URL-адреса, определенного в разделе &lt;Ресурсы&gt; в манифесте.</span><span class="sxs-lookup"><span data-stu-id="c6221-111">The name of a URL resource defined in the &lt;Resources&gt; section of the manifest.</span></span> |

## <a name="child-elements"></a><span data-ttu-id="c6221-112">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c6221-112">Child elements</span></span>

<span data-ttu-id="c6221-113">Нет</span><span class="sxs-lookup"><span data-stu-id="c6221-113">None</span></span>

## <a name="example"></a><span data-ttu-id="c6221-114">Пример</span><span class="sxs-lookup"><span data-stu-id="c6221-114">Example</span></span>

```xml
<SourceLocation resid="pageURL"/>
```
