---
title: Элемент Script в файле манифеста
description: Элемент script определяет параметры скрипта, используемые пользовательской функцией в Excel.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 791f49f15673a029b982e40946f8cc90f02ba887
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608092"
---
# <a name="script-element"></a><span data-ttu-id="73a86-103">Элемент Script</span><span class="sxs-lookup"><span data-stu-id="73a86-103">Script element</span></span>

<span data-ttu-id="73a86-104">Определяет параметры сценариев, используемых пользовательской функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="73a86-104">Defines script settings used by a custom function in Excel.</span></span>

## <a name="attributes"></a><span data-ttu-id="73a86-105">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="73a86-105">Attributes</span></span>

<span data-ttu-id="73a86-106">Нет</span><span class="sxs-lookup"><span data-stu-id="73a86-106">None</span></span>

## <a name="child-elements"></a><span data-ttu-id="73a86-107">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="73a86-107">Child elements</span></span>

|<span data-ttu-id="73a86-108">Элементы</span><span class="sxs-lookup"><span data-stu-id="73a86-108">Elements</span></span>  |  <span data-ttu-id="73a86-109">Обязательный</span><span class="sxs-lookup"><span data-stu-id="73a86-109">Required</span></span>  |  <span data-ttu-id="73a86-110">Описание</span><span class="sxs-lookup"><span data-stu-id="73a86-110">Description</span></span>  |
|:-----|:-----|:-----|
|  [<span data-ttu-id="73a86-111">SourceLocation</span><span class="sxs-lookup"><span data-stu-id="73a86-111">SourceLocation</span></span>](customfunctionssourcelocation.md)  |  <span data-ttu-id="73a86-112">Да</span><span class="sxs-lookup"><span data-stu-id="73a86-112">Yes</span></span>  | <span data-ttu-id="73a86-113">Строка с идентификатором ресурса файла JavaScript, используемого пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="73a86-113">String with resource id of the JavaScript file used by custom functions.</span></span>|

## <a name="example"></a><span data-ttu-id="73a86-114">Пример</span><span class="sxs-lookup"><span data-stu-id="73a86-114">Example</span></span>

```xml
<Script>
    <SourceLocation resid="scriptURL" />
    <!-- The Script element is not used in the Developer Preview. -->
</Script>
```
