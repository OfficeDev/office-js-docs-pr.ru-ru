---
title: Среда выполнения в файле манифеста
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 8fbad8276b3e1d64a6c443cf57d498597d729282
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554001"
---
# <a name="runtime-element"></a><span data-ttu-id="62328-102">Элемент среды выполнения</span><span class="sxs-lookup"><span data-stu-id="62328-102">Runtime element</span></span>

<span data-ttu-id="62328-103">Эта функция доступна предварительная версия.</span><span class="sxs-lookup"><span data-stu-id="62328-103">This feature is in preview.</span></span> <span data-ttu-id="62328-104">Дочерний элемент [`<Runtimes>`](runtimes.md) элемента.</span><span class="sxs-lookup"><span data-stu-id="62328-104">Child element of the [`<Runtimes>`](runtimes.md) element.</span></span> <span data-ttu-id="62328-105">Этот элемент упрощает совместное использование глобальных данных и вызовов функций между пользовательскими функциями Excel и областью задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="62328-105">This element facilitates sharing of global data and function calls between Excel custom functions and the task pane of your add-in.</span></span>

<span data-ttu-id="62328-106">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="62328-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="62328-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="62328-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="contained-in"></a><span data-ttu-id="62328-108">Содержится в</span><span class="sxs-lookup"><span data-stu-id="62328-108">Contained in</span></span>

- [<span data-ttu-id="62328-109">Runtimes</span><span class="sxs-lookup"><span data-stu-id="62328-109">Runtimes</span></span>](runtimes.md)

## <a name="attributes"></a><span data-ttu-id="62328-110">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="62328-110">Attributes</span></span>

|  <span data-ttu-id="62328-111">Атрибут</span><span class="sxs-lookup"><span data-stu-id="62328-111">Attribute</span></span>  |  <span data-ttu-id="62328-112">Обязательный</span><span class="sxs-lookup"><span data-stu-id="62328-112">Required</span></span>  |  <span data-ttu-id="62328-113">Описание</span><span class="sxs-lookup"><span data-stu-id="62328-113">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="62328-114">**время жизни = "Long"**</span><span class="sxs-lookup"><span data-stu-id="62328-114">**lifetime="long"**</span></span>  |  <span data-ttu-id="62328-115">Да</span><span class="sxs-lookup"><span data-stu-id="62328-115">Yes</span></span>  | <span data-ttu-id="62328-116">Всегда должен быть указан как длинное, если вы хотите, чтобы пользовательские функции Excel работали, когда область задач надстройки закрыта.</span><span class="sxs-lookup"><span data-stu-id="62328-116">Should always be listed as long if you want Excel custom functions to work while the task pane of your add-in is closed.</span></span> |
|  <span data-ttu-id="62328-117">**resid**</span><span class="sxs-lookup"><span data-stu-id="62328-117">**resid**</span></span>  |  <span data-ttu-id="62328-118">Да</span><span class="sxs-lookup"><span data-stu-id="62328-118">Yes</span></span>  | <span data-ttu-id="62328-119">Если используется для пользовательских функций Excel, `resid` необходимо указать значение. `TaskPaneAndCustomFunction.Url`</span><span class="sxs-lookup"><span data-stu-id="62328-119">If used for Excel custom functions, the `resid` should point to `TaskPaneAndCustomFunction.Url`.</span></span> |

## <a name="see-also"></a><span data-ttu-id="62328-120">См. также</span><span class="sxs-lookup"><span data-stu-id="62328-120">See also</span></span>

- [<span data-ttu-id="62328-121">Runtimes</span><span class="sxs-lookup"><span data-stu-id="62328-121">Runtimes</span></span>](runtimes.md)
