---
title: Среды выполнения в файле манифеста
description: ''
ms.date: 01/06/2020
localization_priority: Normal
ms.openlocfilehash: ec2b85a92325eb4e36c61f731369ec54d44ef169
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2020
ms.locfileid: "41111179"
---
# <a name="runtimes-element"></a><span data-ttu-id="d51f6-102">Элемент среды выполнения</span><span class="sxs-lookup"><span data-stu-id="d51f6-102">Runtimes element</span></span>

<span data-ttu-id="d51f6-103">Эта функция доступна предварительная версия.</span><span class="sxs-lookup"><span data-stu-id="d51f6-103">This feature is in preview.</span></span> <span data-ttu-id="d51f6-104">Определяет среду выполнения надстройки и позволяет использовать пользовательские функции и область задач для совместного использования глобальных данных и выполнения вызовов функций друг на друга.</span><span class="sxs-lookup"><span data-stu-id="d51f6-104">Specifies the runtime of your add-in and allows custom functions and the task pane to share global data and make function calls into each other.</span></span> <span data-ttu-id="d51f6-105">Должен следовать `<Host>` элементу в файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="d51f6-105">Should follow the `<Host>` element in your manifest file.</span></span>

<span data-ttu-id="d51f6-106">**Тип надстройки:** надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="d51f6-106">**Add-in type:** Task pane</span></span>

## <a name="syntax"></a><span data-ttu-id="d51f6-107">Синтаксис</span><span class="sxs-lookup"><span data-stu-id="d51f6-107">Syntax</span></span>

```XML
<Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
</Runtimes>
```

## <a name="child-elements"></a><span data-ttu-id="d51f6-108">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="d51f6-108">Child elements</span></span>

|  <span data-ttu-id="d51f6-109">Элемент</span><span class="sxs-lookup"><span data-stu-id="d51f6-109">Element</span></span> |  <span data-ttu-id="d51f6-110">Обязательный</span><span class="sxs-lookup"><span data-stu-id="d51f6-110">Required</span></span>  |  <span data-ttu-id="d51f6-111">Описание</span><span class="sxs-lookup"><span data-stu-id="d51f6-111">Description</span></span>  |
|:-----|:-----|:-----|
|  <span data-ttu-id="d51f6-112">**Среда выполнения**</span><span class="sxs-lookup"><span data-stu-id="d51f6-112">**Runtime**</span></span>     | <span data-ttu-id="d51f6-113">Да</span><span class="sxs-lookup"><span data-stu-id="d51f6-113">Yes</span></span> |  <span data-ttu-id="d51f6-114">Среда выполнения надстройки, часто используемая с пользовательскими функциями Excel.</span><span class="sxs-lookup"><span data-stu-id="d51f6-114">The Runtime for your add-in, often used with Excel custom functions.</span></span>

## <a name="see-also"></a><span data-ttu-id="d51f6-115">См. также</span><span class="sxs-lookup"><span data-stu-id="d51f6-115">See also</span></span>

<span data-ttu-id="d51f6-116">-[Сред выполнения](runtimes.md)</span><span class="sxs-lookup"><span data-stu-id="d51f6-116">-[Runtimes](runtimes.md)</span></span>
