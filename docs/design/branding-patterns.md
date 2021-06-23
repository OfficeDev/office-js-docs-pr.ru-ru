---
title: Рекомендации по разработке шаблонов фирменной символики для надстроек Office
description: Узнайте, как маркировать Office надстройку, оставаясь совместимым с визуальным дизайном Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: b42d3a722e4f8805e8c03d2e1a5db528a66f1202
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076373"
---
# <a name="branding-patterns"></a><span data-ttu-id="aeb5b-103">Шаблоны фирменной символики</span><span class="sxs-lookup"><span data-stu-id="aeb5b-103">Branding patterns</span></span>

<span data-ttu-id="aeb5b-104">Эти шаблоны обеспечивают видимость бренда и контекст для пользователей надстройки.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-104">These patterns provide brand visibility and context to your add-in users.</span></span>

## <a name="best-practices"></a><span data-ttu-id="aeb5b-105">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="aeb5b-105">Best practices</span></span>

|<span data-ttu-id="aeb5b-106">Правильно</span><span class="sxs-lookup"><span data-stu-id="aeb5b-106">Do</span></span> |<span data-ttu-id="aeb5b-107">Неправильно</span><span class="sxs-lookup"><span data-stu-id="aeb5b-107">Don't</span></span>|
|:---- |:----|
| <span data-ttu-id="aeb5b-108">Используйте знакомые компоненты пользовательского интерфейса с примененными элементами фирменной символики, такими как оформление и цвет.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-108">Use familiar UI components with applied branding accents like typography and color.</span></span> | <span data-ttu-id="aeb5b-109">Не изобретайте новые компоненты пользовательского интерфейса, которые противоречат установленному интерфейсу Office.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-109">Don't invent new UI components that contradict established Office UI.</span></span> |
| <span data-ttu-id="aeb5b-110">Разместите фирменную символику надстройки в нижнем колонтитуле с панелью фирменной символики внизу пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-110">Place your add-in branding in a brand bar footer at the bottom of your UI.</span></span> | <span data-ttu-id="aeb5b-111">Не повторяйте название области задач в непосредственной близости от панели с фирменной символикой в верхней части пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-111">Don't repeat your task pane name in an immediately adjacent brand bar at the top of your UI.</span></span> |
| <span data-ttu-id="aeb5b-112">Используйте элементы фирменной символики умеренно.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-112">Use brand elements sparingly.</span></span> <span data-ttu-id="aeb5b-113">Разместите свое решение в Office так, чтобы оно было дополняющим.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-113">Fit your solution into Office such that is complementary.</span></span> | <span data-ttu-id="aeb5b-114">Не вставляйте в интерфейс Office слишком много элементов фирменной символики, которые будут отвлекать и путать клиентов.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-114">Don't insert excessively branded elements into Office UI that distract and confuse customers.</span></span> |
| <span data-ttu-id="aeb5b-115">Сделайте свое решение узнаваемым и соедините экраны с помощью единообразных визуальных элементов.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-115">Make your solution recognizable and connect your screens together with consistent visual elements.</span></span> | <span data-ttu-id="aeb5b-116">Не скрывайте свое решение, используя неузнаваемые и непоследовательно применяемые визуальные элементы.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-116">Don't hide your solution with unrecognizable and inconsistently applied visual elements.</span></span> |
| <span data-ttu-id="aeb5b-117">Создайте связь с родительской службой или бизнесом, чтобы клиенты знали и доверяли вашему решению.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-117">Build connection with a parent service or business to ensure that customers know and trust your solution.</span></span> | <span data-ttu-id="aeb5b-118">Не заставляйте клиентов изучать концепцию новой фирменной символики, если есть полезные и понятные связи, которые могут быть использованы для создания доверия и ценности.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-118">Don't make customers learn a new brand concept if there is a useful and understandable relationship that can be leveraged to build trust and value.</span></span> |

<span data-ttu-id="aeb5b-119">Применяйте указанные ниже шаблоны и компоненты, для того чтобы пользователи могли использовать всю полезность вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-119">Apply the following patterns and components as applicable to allow users to embrace the full utility of your add-in.</span></span>

## <a name="brand-bar"></a><span data-ttu-id="aeb5b-120">Панель с фирменной символикой</span><span class="sxs-lookup"><span data-stu-id="aeb5b-120">Brand Bar</span></span>

<span data-ttu-id="aeb5b-121">Панель с фирменной символикой — это место в нижнем колонтитуле, которое содержит фирменное наименование и логотип.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-121">The brand bar is a space in the footer to include your brand name and logo.</span></span> <span data-ttu-id="aeb5b-122">Она также служит ссылкой на ваш фирменный веб-сайт и дополнительным местом доступа.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-122">It also serves as a link to your brand's website and an optional access location.</span></span>

![Бранд-планка, отображаемая в области задач надстройки для Office настольного приложения.](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a><span data-ttu-id="aeb5b-124">Экран-заставка</span><span class="sxs-lookup"><span data-stu-id="aeb5b-124">Splash Screen</span></span>

<span data-ttu-id="aeb5b-125">Используйте этот экран, чтобы отображать вашу фирменную символику, пока надстройка загружается или переходит между состояниями пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="aeb5b-125">Use this screen to display your branding while the add-in is loading or transitioning between UI states.</span></span>

![Экран всплеска бренда, отображающийся в области задач надстройки Office настольного приложения.](../images/add-in-splash-screen.png)
