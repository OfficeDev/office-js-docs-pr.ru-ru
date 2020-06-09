---
title: Элемент Page в файле манифеста
description: Элемент Page определяет параметры страницы HTML, используемые пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: aa8a2807cbf2549ded680a22b17f24513ea76b9a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611500"
---
# <a name="page-element"></a><span data-ttu-id="906f5-103">Элемент Page</span><span class="sxs-lookup"><span data-stu-id="906f5-103">Page element</span></span>

<span data-ttu-id="906f5-104">Определяет параметры HTML-страницы, используемые пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="906f5-104">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="906f5-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="906f5-105">Attributes</span></span>

<span data-ttu-id="906f5-106">Нет</span><span class="sxs-lookup"><span data-stu-id="906f5-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="906f5-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="906f5-107">Child elements</span></span>

|  <span data-ttu-id="906f5-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="906f5-108">Element</span></span>  |  <span data-ttu-id="906f5-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="906f5-109">Required</span></span>  |  <span data-ttu-id="906f5-110">Описание</span><span class="sxs-lookup"><span data-stu-id="906f5-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="906f5-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="906f5-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="906f5-112">Да</span><span class="sxs-lookup"><span data-stu-id="906f5-112">Yes</span></span>  | <span data-ttu-id="906f5-113">Строка с идентификатором ресурса HTML-файла, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="906f5-113">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="906f5-114">Пример</span><span class="sxs-lookup"><span data-stu-id="906f5-114">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
