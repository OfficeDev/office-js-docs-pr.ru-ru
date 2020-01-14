---
title: Среда выполнения в файле манифеста
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: 68def44ba74733934198ac3b32fa1fe649156766
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111172"
---
# <a name="runtime-element"></a><span data-ttu-id="0479d-102">Элемент среды выполнения</span><span class="sxs-lookup"><span data-stu-id="0479d-102">Runtime element</span></span>

<span data-ttu-id="0479d-103">Эта функция доступна предварительная версия.</span><span class="sxs-lookup"><span data-stu-id="0479d-103">This feature is in preview.</span></span> <span data-ttu-id="0479d-104">Дочерний элемент [`<Runtimes>`](runtime.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="0479d-104">Child element of the [`<Runtimes>`](runtime.md) element.</span></span> <span data-ttu-id="0479d-105">Этот элемент упрощает совместное использование глобальных данных и вызовов функций между пользовательскими функциями Excel и областью задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="0479d-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span> 

## <a name="contained-in"></a><span data-ttu-id="0479d-106">Содержится в</span><span class="sxs-lookup"><span data-stu-id="0479d-106">Contained in</span></span>

<span data-ttu-id="0479d-107">-[Сред выполнения](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="0479d-107">-[Runtimes](runtimes.md)</span></span>

<span data-ttu-id="0479d-108">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="0479d-108">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="0479d-109">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="0479d-109">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="attributes"></a><span data-ttu-id="0479d-110">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="0479d-110">Attributes</span></span>

|  <span data-ttu-id="0479d-111">Атрибут</span><span class="sxs-lookup"><span data-stu-id="0479d-111">Attribute</span></span>  |  <span data-ttu-id="0479d-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0479d-112">Required</span></span>  |  <span data-ttu-id="0479d-113">Описание</span><span class="sxs-lookup"><span data-stu-id="0479d-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="0479d-114">**время жизни = "Long"**</span><span class="sxs-lookup"><span data-stu-id="0479d-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="0479d-115">Да</span><span class="sxs-lookup"><span data-stu-id="0479d-115">Yes</span></span>  | <span data-ttu-id="0479d-116">Всегда должен быть указан как длинное, если вы хотите, чтобы пользовательские функции Excel работали, когда область задач надстройки закрыта.</span><span class="sxs-lookup"><span data-stu-id="0479d-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="0479d-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="0479d-117">**resid**</span></span>  |  <span data-ttu-id="0479d-118">Да</span><span class="sxs-lookup"><span data-stu-id="0479d-118">Yes</span></span>  | <span data-ttu-id="0479d-119">Если используется для пользовательских функций Excel, `resid` необходимо указать значение. `TaskPaneAndCustomFunction.Url`</span><span class="sxs-lookup"><span data-stu-id="0479d-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="0479d-120">См. также</span><span class="sxs-lookup"><span data-stu-id="0479d-120">See also</span></span>

<span data-ttu-id="0479d-121">-[Полняющего](runtime.md)</span><span class="sxs-lookup"><span data-stu-id="0479d-121">-[Runtime](runtime.md)</span></span>
