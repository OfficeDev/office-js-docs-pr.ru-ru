---
title: Среды выполнения в файле манифеста
description: ''
ms.date: 01/24/2020
localization_priority: Normal
ms.openlocfilehash: 6682887935ee6894b5a311ad519408067452bb23
ms.sourcegitcommit: 72d719165cc2b64ac9d3c51fb8be277dfde7d2eb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/25/2020
ms.locfileid: "41554008"
---
# <a name="runtimes-element"></a><span data-ttu-id="a961a-102">Элемент среды выполнения</span><span class="sxs-lookup"><span data-stu-id="a961a-102">Runtimes element</span></span>

<span data-ttu-id="a961a-103">Эта функция доступна предварительная версия.</span><span class="sxs-lookup"><span data-stu-id="a961a-103">This feature is in preview.</span></span> <span data-ttu-id="a961a-104">Определяет среду выполнения надстройки и позволяет использовать пользовательские функции и область задач для совместного использования глобальных данных и выполнения вызовов функций друг на друга.</span><span class="sxs-lookup"><span data-stu-id="a961a-104">Specifies the runtime of your add-in and allows custom functions and the task pane to share global data and make function calls into each other.</span></span> <span data-ttu-id="a961a-105">Должен следовать `<Host>` элементу в файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="a961a-105">Should follow the `<Host>` element in your manifest file.</span></span>

<span data-ttu-id="a961a-106">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="a961a-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="a961a-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="a961a-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a><span data-ttu-id="a961a-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="a961a-108">Child elements</span></span>

|  <span data-ttu-id="a961a-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="a961a-109">Element</span></span> |  <span data-ttu-id="a961a-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="a961a-110">Required</span></span>  |  <span data-ttu-id="a961a-111">Описание</span><span class="sxs-lookup"><span data-stu-id="a961a-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="a961a-112">**Среда выполнения**</span><span class="sxs-lookup"><span data-stu-id="a961a-112">**Runtime**</span></span>     | <span data-ttu-id="a961a-113">Да</span><span class="sxs-lookup"><span data-stu-id="a961a-113">Yes</span></span> |  <span data-ttu-id="a961a-114">Среда выполнения надстройки, часто используемая с пользовательскими функциями Excel.</span><span class="sxs-lookup"><span data-stu-id="a961a-114">The Runtime for your add-in, often used with Excel custom functions.</span></span>

## <a name="see-also"></a><span data-ttu-id="a961a-115">См. также</span><span class="sxs-lookup"><span data-stu-id="a961a-115">See also</span></span>

- [<span data-ttu-id="a961a-116">Среда выполнения</span><span class="sxs-lookup"><span data-stu-id="a961a-116">Runtime</span></span>](runtime.md)
