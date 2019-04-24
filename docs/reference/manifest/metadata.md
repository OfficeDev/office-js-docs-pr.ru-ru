---
title: Элемент Metadata в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: a3aecb1983905658f3a55fdb8bf0629a8d5ef474
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452048"
---
# <a name="metadata-element"></a><span data-ttu-id="c07ef-102">Элемент Metadata</span><span class="sxs-lookup"><span data-stu-id="c07ef-102">Metadata element</span></span>

<span data-ttu-id="c07ef-103">Определяет параметры метаданных, используемые пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="c07ef-103">Defines the metadata settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="c07ef-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="c07ef-104">Attributes</span></span>

<span data-ttu-id="c07ef-105">Нет</span><span class="sxs-lookup"><span data-stu-id="c07ef-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="c07ef-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="c07ef-106">Child elements</span></span>

|  <span data-ttu-id="c07ef-107">Элемент</span><span class="sxs-lookup"><span data-stu-id="c07ef-107">Element</span></span>  |  <span data-ttu-id="c07ef-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="c07ef-108">Required</span></span>  |  <span data-ttu-id="c07ef-109">Описание</span><span class="sxs-lookup"><span data-stu-id="c07ef-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="c07ef-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="c07ef-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="c07ef-111">Да</span><span class="sxs-lookup"><span data-stu-id="c07ef-111">Yes</span></span>  | <span data-ttu-id="c07ef-112">Строка с идентификатором ресурса JSON-файла, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="c07ef-112">String with the resource id of the JSON file used by custom functions.</span></span> |

## <a name="example"></a><span data-ttu-id="c07ef-113">Пример</span><span class="sxs-lookup"><span data-stu-id="c07ef-113">Example</span></span>

```xml
<Metadata>
    <SourceLocation resid="JSON-URL" />
</Metadata>
```
