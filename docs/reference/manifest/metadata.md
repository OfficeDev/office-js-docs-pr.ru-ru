---
title: Элемент Metadata в файле манифеста
description: Элемент Metadata определяет параметры метаданных, используемые пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8ea81818aa96b407ce386ec318495ec5ba773d05
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718071"
---
# <a name="metadata-element"></a><span data-ttu-id="96aab-103">Элемент Metadata</span><span class="sxs-lookup"><span data-stu-id="96aab-103">Metadata element</span></span>

<span data-ttu-id="96aab-104">Определяет параметры метаданных, используемые пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="96aab-104">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="96aab-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="96aab-105">Attributes</span></span>

<span data-ttu-id="96aab-106">Нет</span><span class="sxs-lookup"><span data-stu-id="96aab-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="96aab-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="96aab-107">Child elements</span></span>

|  <span data-ttu-id="96aab-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="96aab-108">Element</span></span>  |  <span data-ttu-id="96aab-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="96aab-109">Required</span></span>  |  <span data-ttu-id="96aab-110">Описание</span><span class="sxs-lookup"><span data-stu-id="96aab-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="96aab-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="96aab-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="96aab-112">Да</span><span class="sxs-lookup"><span data-stu-id="96aab-112">Yes</span></span>  | <span data-ttu-id="96aab-113">Строка с идентификатором ресурса JSON-файла, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="96aab-113">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="96aab-114">Пример</span><span class="sxs-lookup"><span data-stu-id="96aab-114">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
