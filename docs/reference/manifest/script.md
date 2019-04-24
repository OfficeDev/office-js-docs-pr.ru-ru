---
title: Элемент Script в файле манифеста
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8352ada0eeb6af071d5f20f750dcdeaefe31e918
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450438"
---
# <a name="script-element"></a><span data-ttu-id="89f99-102">Элемент Script</span><span class="sxs-lookup"><span data-stu-id="89f99-102">Script element</span></span>

<span data-ttu-id="89f99-103">Определяет параметры сценариев, используемых пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="89f99-103">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="89f99-104">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="89f99-104">Attributes</span></span>

<span data-ttu-id="89f99-105">Нет</span><span class="sxs-lookup"><span data-stu-id="89f99-105">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="89f99-106">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="89f99-106">Child elements</span></span>

|<span data-ttu-id="89f99-107">Элементы</span><span class="sxs-lookup"><span data-stu-id="89f99-107">Elements</span></span>  |  <span data-ttu-id="89f99-108">Обязательный</span><span class="sxs-lookup"><span data-stu-id="89f99-108">Required</span></span>  |  <span data-ttu-id="89f99-109">Описание</span><span class="sxs-lookup"><span data-stu-id="89f99-109">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="89f99-110">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="89f99-110">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="89f99-111">Да</span><span class="sxs-lookup"><span data-stu-id="89f99-111">Yes</span></span>  | <span data-ttu-id="89f99-112">Строка с идентификатором ресурса файла JavaScript, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="89f99-112">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="89f99-113">Пример</span><span class="sxs-lookup"><span data-stu-id="89f99-113">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
