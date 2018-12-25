---
title: Элемент Metadata в файле манифеста
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 79038fc13eba76176be19e484ffa57e64727bf94
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432664"
---
# <a name="metadata-element"></a><span data-ttu-id="bd30a-102">Элемент Metadata</span><span class="sxs-lookup"><span data-stu-id="bd30a-102">MetaData element</span></span>

<span data-ttu-id="bd30a-103">Определяет параметры метаданных, используемые пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="bd30a-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="bd30a-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="bd30a-104">Attributes</span></span>

<span data-ttu-id="bd30a-105">Нет</span><span class="sxs-lookup"><span data-stu-id="bd30a-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="bd30a-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="bd30a-106">Child elements</span></span>

|  <span data-ttu-id="bd30a-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="bd30a-107">Element</span></span>  |  <span data-ttu-id="bd30a-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="bd30a-108">Required</span></span>  |  <span data-ttu-id="bd30a-109">Описание</span><span class="sxs-lookup"><span data-stu-id="bd30a-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="bd30a-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="bd30a-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="bd30a-111">Да</span><span class="sxs-lookup"><span data-stu-id="bd30a-111">Yes</span></span>  | <span data-ttu-id="bd30a-112">Строка с идентификатором ресурса JSON-файла, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="bd30a-112">String with the resource id of the HTML file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="bd30a-113">Пример</span><span class="sxs-lookup"><span data-stu-id="bd30a-113">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
