---
title: Элемент Script в файле манифеста
description: Элемент script определяет параметры скрипта, используемые пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: f05fc85bd0454c340f4352bb73f299b9e7730224
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720416"
---
# <a name="script-element"></a><span data-ttu-id="4a41a-103">Элемент Script</span><span class="sxs-lookup"><span data-stu-id="4a41a-103">Script element</span></span>

<span data-ttu-id="4a41a-104">Определяет параметры сценариев, используемых пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="4a41a-104">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="4a41a-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4a41a-105">Attributes</span></span>

<span data-ttu-id="4a41a-106">Нет</span><span class="sxs-lookup"><span data-stu-id="4a41a-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="4a41a-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="4a41a-107">Child elements</span></span>

|<span data-ttu-id="4a41a-108">Элементы</span><span class="sxs-lookup"><span data-stu-id="4a41a-108">Elements</span></span>  |  <span data-ttu-id="4a41a-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="4a41a-109">Required</span></span>  |  <span data-ttu-id="4a41a-110">Описание</span><span class="sxs-lookup"><span data-stu-id="4a41a-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="4a41a-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="4a41a-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="4a41a-112">Да</span><span class="sxs-lookup"><span data-stu-id="4a41a-112">Yes</span></span>  | <span data-ttu-id="4a41a-113">Строка с идентификатором ресурса файла JavaScript, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="4a41a-113">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="4a41a-114">Пример</span><span class="sxs-lookup"><span data-stu-id="4a41a-114">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
