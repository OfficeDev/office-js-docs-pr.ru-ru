---
title: Элемент Page в файле манифеста
description: Элемент Page определяет параметры страницы HTML, используемые пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 0c56b955b79f9052ee2c89a391dd95b2975d69c2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720486"
---
# <a name="page-element"></a><span data-ttu-id="3edf0-103">Элемент Page</span><span class="sxs-lookup"><span data-stu-id="3edf0-103">Page element</span></span>

<span data-ttu-id="3edf0-104">Определяет параметры HTML-страницы, используемые пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="3edf0-104">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="3edf0-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3edf0-105">Attributes</span></span>

<span data-ttu-id="3edf0-106">Нет</span><span class="sxs-lookup"><span data-stu-id="3edf0-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="3edf0-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="3edf0-107">Child elements</span></span>

|  <span data-ttu-id="3edf0-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="3edf0-108">Element</span></span>  |  <span data-ttu-id="3edf0-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="3edf0-109">Required</span></span>  |  <span data-ttu-id="3edf0-110">Описание</span><span class="sxs-lookup"><span data-stu-id="3edf0-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="3edf0-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="3edf0-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="3edf0-112">Да</span><span class="sxs-lookup"><span data-stu-id="3edf0-112">Yes</span></span>  | <span data-ttu-id="3edf0-113">Строка с идентификатором ресурса HTML-файла, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="3edf0-113">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="3edf0-114">Пример</span><span class="sxs-lookup"><span data-stu-id="3edf0-114">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
