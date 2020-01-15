---
title: Среда выполнения в файле манифеста
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 945a30527632b23a594d7bfb82cec94e74754249
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120637"
---
# <a name="runtime-element"></a><span data-ttu-id="a8617-102">Элемент среды выполнения</span><span class="sxs-lookup"><span data-stu-id="a8617-102">Runtime element</span></span>

<span data-ttu-id="a8617-103">Эта функция доступна предварительная версия.</span><span class="sxs-lookup"><span data-stu-id="a8617-103">This feature is in preview.</span></span> <span data-ttu-id="a8617-104">Дочерний элемент [`<Runtimes>`](runtime.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="a8617-104">Child element of the [`<Runtimes>`](runtime.md) element.</span></span> <span data-ttu-id="a8617-105">Этот элемент упрощает совместное использование глобальных данных и вызовов функций между пользовательскими функциями Excel и областью задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="a8617-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span>

<span data-ttu-id="a8617-106">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="a8617-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="a8617-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="a8617-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="a8617-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="a8617-108">Contained in</span></span>

<span data-ttu-id="a8617-109">-[Сред выполнения](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="a8617-109">-[Runtimes](runtimes.md)</span></span>

## <a name="attributes"></a><span data-ttu-id="a8617-110">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="a8617-110">Attributes</span></span>

|  <span data-ttu-id="a8617-111">Атрибут</span><span class="sxs-lookup"><span data-stu-id="a8617-111">Attribute</span></span>  |  <span data-ttu-id="a8617-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a8617-112">Required</span></span>  |  <span data-ttu-id="a8617-113">Описание</span><span class="sxs-lookup"><span data-stu-id="a8617-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a8617-114">**время жизни = "Long"**</span><span class="sxs-lookup"><span data-stu-id="a8617-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="a8617-115">Да</span><span class="sxs-lookup"><span data-stu-id="a8617-115">Yes</span></span>  | <span data-ttu-id="a8617-116">Всегда должен быть указан как длинное, если вы хотите, чтобы пользовательские функции Excel работали, когда область задач надстройки закрыта.</span><span class="sxs-lookup"><span data-stu-id="a8617-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="a8617-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="a8617-117">**resid**</span></span>  |  <span data-ttu-id="a8617-118">Да</span><span class="sxs-lookup"><span data-stu-id="a8617-118">Yes</span></span>  | <span data-ttu-id="a8617-119">Если используется для пользовательских функций Excel, `resid` необходимо указать значение. `TaskPaneAndCustomFunction.Url`</span><span class="sxs-lookup"><span data-stu-id="a8617-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="a8617-120">См. также</span><span class="sxs-lookup"><span data-stu-id="a8617-120">See also</span></span>

<span data-ttu-id="a8617-121">-[Полняющего](runtime.md)</span><span class="sxs-lookup"><span data-stu-id="a8617-121">-[Runtime](runtime.md)</span></span>
