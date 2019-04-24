---
title: Элемент Page в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f85cc3a834f628a7390f3b96faa596145c7d331a
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452076"
---
# <a name="page-element"></a><span data-ttu-id="40c0d-102">Элемент Page</span><span class="sxs-lookup"><span data-stu-id="40c0d-102">Page element</span></span>

<span data-ttu-id="40c0d-103">Определяет параметры HTML-страницы, используемые пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="40c0d-103">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="40c0d-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="40c0d-104">Attributes</span></span>

<span data-ttu-id="40c0d-105">Нет</span><span class="sxs-lookup"><span data-stu-id="40c0d-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="40c0d-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="40c0d-106">Child elements</span></span>

|  <span data-ttu-id="40c0d-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="40c0d-107">Element</span></span>  |  <span data-ttu-id="40c0d-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="40c0d-108">Required</span></span>  |  <span data-ttu-id="40c0d-109">Описание</span><span class="sxs-lookup"><span data-stu-id="40c0d-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="40c0d-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="40c0d-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="40c0d-111">Да</span><span class="sxs-lookup"><span data-stu-id="40c0d-111">Yes</span></span>  | <span data-ttu-id="40c0d-112">Строка с идентификатором ресурса HTML-файла, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="40c0d-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="40c0d-113">Пример</span><span class="sxs-lookup"><span data-stu-id="40c0d-113">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
