---
title: Элемент Page в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 83bafd24d0b56322ea5f7d51025f2416be019168
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433735"
---
# <a name="page-element"></a><span data-ttu-id="fec46-102">Элемент Page</span><span class="sxs-lookup"><span data-stu-id="fec46-102">Page element</span></span>

<span data-ttu-id="fec46-103">Определяет параметры HTML-страницы, используемые пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="fec46-103">Defines HTML page settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="fec46-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="fec46-104">Attributes</span></span>

<span data-ttu-id="fec46-105">Нет</span><span class="sxs-lookup"><span data-stu-id="fec46-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="fec46-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="fec46-106">Child elements</span></span>

|  <span data-ttu-id="fec46-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="fec46-107">Element</span></span>  |  <span data-ttu-id="fec46-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="fec46-108">Required</span></span>  |  <span data-ttu-id="fec46-109">Описание</span><span class="sxs-lookup"><span data-stu-id="fec46-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="fec46-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="fec46-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="fec46-111">Да</span><span class="sxs-lookup"><span data-stu-id="fec46-111">Yes</span></span>  | <span data-ttu-id="fec46-112">Строка с идентификатором ресурса HTML-файла, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="fec46-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="fec46-113">Пример</span><span class="sxs-lookup"><span data-stu-id="fec46-113">Example</span></span>

```xml
<Page>
    <SourceLocation resid="pageURL"/>
</Page>
```
