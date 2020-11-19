---
title: Рекомендации по разработке шаблонов фирменной символики для надстроек Office
description: Узнайте, как настроить фирменную символику для надстройки Office, чтобы обеспечить совместимость с визуальным дизайном Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 9571b2771767a60f715ac7f3bc83a501907039d9
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132090"
---
# <a name="branding-patterns"></a><span data-ttu-id="ab0e4-103">Шаблоны фирменной символики</span><span class="sxs-lookup"><span data-stu-id="ab0e4-103">Branding patterns</span></span>

<span data-ttu-id="ab0e4-104">Эти шаблоны обеспечивают видимость и контекст фирменного стиля для пользователей надстроек.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-104">These patterns provide brand visibility and context to your add-in users.</span></span>

## <a name="best-practices"></a><span data-ttu-id="ab0e4-105">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="ab0e4-105">Best practices</span></span>

|<span data-ttu-id="ab0e4-106">Правильно</span><span class="sxs-lookup"><span data-stu-id="ab0e4-106">Do</span></span> |<span data-ttu-id="ab0e4-107">Неправильно</span><span class="sxs-lookup"><span data-stu-id="ab0e4-107">Don't</span></span>|
|:---- |:----|
| <span data-ttu-id="ab0e4-108">Используйте знакомые компоненты пользовательского интерфейса с примененными элементами фирменной символики, такими как оформление и цвет.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-108">Use familiar UI components with applied branding accents like typography and color.</span></span> | <span data-ttu-id="ab0e4-109">Не изобретайте новые компоненты пользовательского интерфейса, которые противоречат установленному интерфейсу Office.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-109">Don't invent new UI components that contradict established Office UI.</span></span> |
| <span data-ttu-id="ab0e4-110">Разместите фирменную символику надстройки в нижнем колонтитуле с панелью фирменной символики внизу пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-110">Place your add-in branding in a brand bar footer at the bottom of your UI.</span></span> | <span data-ttu-id="ab0e4-111">Не повторяйте название области задач в непосредственной близости от панели с фирменной символикой в верхней части пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-111">Don't repeat your task pane name in an immediately adjacent brand bar at the top of your UI.</span></span> |
| <span data-ttu-id="ab0e4-112">Используйте элементы фирменной символики умеренно.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-112">Use brand elements sparingly.</span></span> <span data-ttu-id="ab0e4-113">Разместите свое решение в Office так, чтобы оно было дополняющим.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-113">Fit your solution into Office such that is complementary.</span></span> | <span data-ttu-id="ab0e4-114">Не вставляйте в интерфейс Office слишком много элементов фирменной символики, которые будут отвлекать и путать клиентов.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-114">Don't insert excessively branded elements into Office UI that distract and confuse customers.</span></span> |
| <span data-ttu-id="ab0e4-115">Сделайте свое решение узнаваемым и соедините экраны с помощью единообразных визуальных элементов.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-115">Make your solution recognizable and connect your screens together with consistent visual elements.</span></span> | <span data-ttu-id="ab0e4-116">Не скрывайте свое решение, используя неузнаваемые и непоследовательно применяемые визуальные элементы.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-116">Don't hide your solution with unrecognizable and inconsistently applied visual elements.</span></span> |
| <span data-ttu-id="ab0e4-117">Создайте связь с родительской службой или бизнесом, чтобы клиенты знали и доверяли вашему решению.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-117">Build connection with a parent service or business to ensure that customers know and trust your solution.</span></span> | <span data-ttu-id="ab0e4-118">Не заставляйте клиентов изучать концепцию новой фирменной символики, если есть полезные и понятные связи, которые могут быть использованы для создания доверия и ценности.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-118">Don't make customers learn a new brand concept if there is a useful and understandable relationship that can be leveraged to build trust and value.</span></span> |

<span data-ttu-id="ab0e4-119">Применяйте указанные ниже шаблоны и компоненты, для того чтобы пользователи могли использовать всю полезность вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-119">Apply the following patterns and components as applicable to allow users to embrace the full utility of your add-in.</span></span>

## <a name="brand-bar"></a><span data-ttu-id="ab0e4-120">Панель с фирменной символикой</span><span class="sxs-lookup"><span data-stu-id="ab0e4-120">Brand Bar</span></span>

<span data-ttu-id="ab0e4-121">Панель с фирменной символикой — это место в нижнем колонтитуле, которое содержит фирменное наименование и логотип.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-121">The brand bar is a space in the footer to include your brand name and logo.</span></span> <span data-ttu-id="ab0e4-122">Она также служит ссылкой на ваш фирменный веб-сайт и дополнительным местом доступа.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-122">It also serves as a link to your brand's website and an optional access location.</span></span>

![Панель фирменной символики, отображаемая в области задач надстройки для настольного приложения Office](../images/add-in-brand-bar.png)

## <a name="splash-screen"></a><span data-ttu-id="ab0e4-124">Экран-заставка</span><span class="sxs-lookup"><span data-stu-id="ab0e4-124">Splash Screen</span></span>

<span data-ttu-id="ab0e4-125">Используйте этот экран, чтобы отображать вашу фирменную символику, пока надстройка загружается или переходит между состояниями пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="ab0e4-125">Use this screen to display your branding while the add-in is loading or transitioning between UI states.</span></span>

![Экран-заставка фирменного стиля, отображаемый в области задач надстройки для настольного приложения Office](../images/add-in-splash-screen.png)
