---
title: Элемент Metadata в файле манифеста
description: Элемент Metadata определяет параметры метаданных, используемые пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 01be124b5526ce8328e0a20b8ff7d21ba6da96bc
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611766"
---
# <a name="metadata-element"></a><span data-ttu-id="01c37-103">Элемент Metadata</span><span class="sxs-lookup"><span data-stu-id="01c37-103">Metadata element</span></span>

<span data-ttu-id="01c37-104">Определяет параметры метаданных, используемые пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="01c37-104">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="01c37-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="01c37-105">Attributes</span></span>

<span data-ttu-id="01c37-106">Нет</span><span class="sxs-lookup"><span data-stu-id="01c37-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="01c37-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="01c37-107">Child elements</span></span>

|  <span data-ttu-id="01c37-108">Элемент</span><span class="sxs-lookup"><span data-stu-id="01c37-108">Element</span></span>  |  <span data-ttu-id="01c37-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="01c37-109">Required</span></span>  |  <span data-ttu-id="01c37-110">Описание</span><span class="sxs-lookup"><span data-stu-id="01c37-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="01c37-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="01c37-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="01c37-112">Да</span><span class="sxs-lookup"><span data-stu-id="01c37-112">Yes</span></span>  | <span data-ttu-id="01c37-113">Строка с идентификатором ресурса JSON-файла, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="01c37-113">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="01c37-114">Пример</span><span class="sxs-lookup"><span data-stu-id="01c37-114">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
