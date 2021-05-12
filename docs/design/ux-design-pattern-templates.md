---
title: Конструктивные шаблоны пользовательского интерфейса для надстроек Office
description: Получите обзор шаблонов проектирования пользовательского интерфейса для надстройок Office, включая шаблоны для навигации, проверки подлинности, первого запуска и брендинга.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 8544b56b85a25d522c95546b42a78fe01a3c2586
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330110"
---
# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="d5a77-103">Конструктивные шаблоны пользовательского интерфейса для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d5a77-103">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="d5a77-104">Проектирование пользовательского интерфейса для надстроек Office должно обеспечивать удобство работы для пользователей Office и расширять функциональный набор Office благодаря плавной интеграции в интерфейс Office по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="d5a77-104">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="d5a77-105">Наши шаблоны пользовательского интерфейса состоят из компонентов.</span><span class="sxs-lookup"><span data-stu-id="d5a77-105">Our UX patterns are composed of components.</span></span> <span data-ttu-id="d5a77-106">Компоненты — это элементы управления, которые помогают клиентам взаимодействовать с элементами программного обеспечения или службы.</span><span class="sxs-lookup"><span data-stu-id="d5a77-106">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="d5a77-107">Кнопки, элементы навигации и меню — это примеры общих компонентов, которые часто отличаются единым стилем и поведением.</span><span class="sxs-lookup"><span data-stu-id="d5a77-107">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="d5a77-108">[Компоненты пользовательского интерфейса React](using-office-ui-fabric-react.md) и ведут себя как часть Office, как и нейтральные в рамках компоненты [Office UI Fabric JS](fabric-core.md).</span><span class="sxs-lookup"><span data-stu-id="d5a77-108">[Fluent UI React components](using-office-ui-fabric-react.md) look and behave like a part of Office, as do the framework-neutral components of [Office UI Fabric JS](fabric-core.md).</span></span> <span data-ttu-id="d5a77-109">Для интеграции с Office.</span><span class="sxs-lookup"><span data-stu-id="d5a77-109">Take advantage of either set of components to integrate with Office.</span></span> <span data-ttu-id="d5a77-110">Кроме того, если у вашей надстройки есть свой язык предварительного компоненты, его не нужно отбрасывать.</span><span class="sxs-lookup"><span data-stu-id="d5a77-110">Alternatively, if your add-in has its own preexisting component language, you don't need to discard it.</span></span> <span data-ttu-id="d5a77-111">Найдите возможности сохранить его, интегрируя надстройку с Office.</span><span class="sxs-lookup"><span data-stu-id="d5a77-111">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="d5a77-112">Рассмотрите способы изменения стилистических элементов и удаления конфликтов или примените понятные пользователям стили и поведение.</span><span class="sxs-lookup"><span data-stu-id="d5a77-112">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="d5a77-113">Предоставленные шаблоны — это наилучшие решения, основанные на общих сценариях клиентов и исследованиях работы пользователей.</span><span class="sxs-lookup"><span data-stu-id="d5a77-113">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="d5a77-114">Они предназначены для обеспечения быстрой точки входа для разработки и разработки надстройок, а также руководства по достижению баланса между элементами бренда Майкрософт и вашими собственными.</span><span class="sxs-lookup"><span data-stu-id="d5a77-114">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft brand elements and your own.</span></span> <span data-ttu-id="d5a77-115">Предоставление чистого и современного пользовательского интерфейса, которое уравновешивало элементы дизайна с помощью языка разработки пользовательского интерфейса Fluent Корпорации Майкрософт и уникальной фирменной идентичности партнера, может помочь увеличить удержание пользователей и принятие вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="d5a77-115">Providing a clean, modern user experience that balances design elements from Microsoft's Fluent UI design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="d5a77-116">Используйте шаблонные заготовки пользовательского интерфейса для того, чтобы:</span><span class="sxs-lookup"><span data-stu-id="d5a77-116">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="d5a77-117">применять решения в распространенных клиентских сценариях;</span><span class="sxs-lookup"><span data-stu-id="d5a77-117">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="d5a77-118">следовать рекомендациям по оформлению;</span><span class="sxs-lookup"><span data-stu-id="d5a77-118">Apply design best practices.</span></span>
* <span data-ttu-id="d5a77-119">Включение компонентов и стилей пользовательского интерфейса [Fluent.](https://developer.microsoft.com/fluentui#/get-started)</span><span class="sxs-lookup"><span data-stu-id="d5a77-119">Incorporate [Fluent UI](https://developer.microsoft.com/fluentui#/get-started) components and styles.</span></span>
* <span data-ttu-id="d5a77-120">создавать надстройки, внешний вид которых согласован со стандартным пользовательским интерфейсом Office;</span><span class="sxs-lookup"><span data-stu-id="d5a77-120">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="d5a77-121">формировать и визуализировать пользовательский интерфейс.</span><span class="sxs-lookup"><span data-stu-id="d5a77-121">Ideate and visualize UX.</span></span>

## <a name="getting-started"></a><span data-ttu-id="d5a77-122">Начало работы</span><span class="sxs-lookup"><span data-stu-id="d5a77-122">Getting started</span></span>

<span data-ttu-id="d5a77-123">Шаблоны организованы по ключевым действиям или функциональным возможностям, которые часто используются в надстройке.</span><span class="sxs-lookup"><span data-stu-id="d5a77-123">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="d5a77-124">Основные группы:</span><span class="sxs-lookup"><span data-stu-id="d5a77-124">The main groups are:</span></span>

* [<span data-ttu-id="d5a77-125">Первый запуск (FRE)</span><span class="sxs-lookup"><span data-stu-id="d5a77-125">First run experience (FRE)</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="d5a77-126">Проверка подлинности</span><span class="sxs-lookup"><span data-stu-id="d5a77-126">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="d5a77-127">Навигация</span><span class="sxs-lookup"><span data-stu-id="d5a77-127">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="d5a77-128">Разработка фирменной символики</span><span class="sxs-lookup"><span data-stu-id="d5a77-128">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="d5a77-129">Просмотрите каждую группу, чтобы получить представление о том, как можно создавать свои надстройки с использованием рекомендаций.</span><span class="sxs-lookup"><span data-stu-id="d5a77-129">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>

> [!NOTE]
> <span data-ttu-id="d5a77-130">Примеры экранов, демонстрируемые в этой документации, созданы и отображены с разрешением **1366x768**.</span><span class="sxs-lookup"><span data-stu-id="d5a77-130">The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**.</span></span>

## <a name="see-also"></a><span data-ttu-id="d5a77-131">См. также</span><span class="sxs-lookup"><span data-stu-id="d5a77-131">See also</span></span>

* [<span data-ttu-id="d5a77-132">Наборы инструментов для разработки</span><span class="sxs-lookup"><span data-stu-id="d5a77-132">Design tool kits</span></span>](design-toolkits.md)
* [<span data-ttu-id="d5a77-133">Пользовательский интерфейс Fluent</span><span class="sxs-lookup"><span data-stu-id="d5a77-133">Fluent UI</span></span>](https://developer.microsoft.com/fluentui#)
* [<span data-ttu-id="d5a77-134">Рекомендации по разработке надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d5a77-134">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
* [<span data-ttu-id="d5a77-135">Fluent UI React в Office надстройки</span><span class="sxs-lookup"><span data-stu-id="d5a77-135">Fluent UI React in Office Add-ins</span></span>](using-office-ui-fabric-react.md)
