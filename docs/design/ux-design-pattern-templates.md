---
title: Конструктивные шаблоны пользовательского интерфейса для надстроек Office
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 40b36fb138169bdf848e5f58569e6fc3dee8c09b
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449011"
---
# <a name="ux-design-patterns-for-office-add-ins"></a><span data-ttu-id="413ff-102">Конструктивные шаблоны пользовательского интерфейса для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="413ff-102">UX design patterns for Office Add-ins</span></span>

<span data-ttu-id="413ff-103">Проектирование пользовательского интерфейса для надстроек Office должно обеспечивать удобство работы для пользователей Office и расширять функциональный набор Office благодаря плавной интеграции в интерфейс Office по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="413ff-103">Designing the user experience for Office Add-ins should provide a compelling experience for Office users and extend the overall Office experience by fitting seamlessly within the default Office UI.</span></span>  

<span data-ttu-id="413ff-104">Наши шаблоны пользовательского интерфейса состоят из компонентов.</span><span class="sxs-lookup"><span data-stu-id="413ff-104">Our UX patterns are composed of components.</span></span> <span data-ttu-id="413ff-105">Компоненты — это элементы управления, которые помогают клиентам взаимодействовать с элементами программного обеспечения или службы.</span><span class="sxs-lookup"><span data-stu-id="413ff-105">Components are controls that help your customers interact with elements of your software or service.</span></span> <span data-ttu-id="413ff-106">Кнопки, элементы навигации и меню — это примеры общих компонентов, которые часто отличаются единым стилем и поведением.</span><span class="sxs-lookup"><span data-stu-id="413ff-106">Buttons, navigation, and menus are examples of common components that often have consistent styles and behaviors.</span></span>

<span data-ttu-id="413ff-107">Office UI Fabric обрабатывает компоненты, обеспечивая их полную совместимость с Office.</span><span class="sxs-lookup"><span data-stu-id="413ff-107">Office UI Fabric renders components that look and behave like a part of Office.</span></span> <span data-ttu-id="413ff-108">Воспользуйтесь преимуществами Fabric для легкой интеграции с Office.</span><span class="sxs-lookup"><span data-stu-id="413ff-108">Take advantage of Fabric to easily integrate with Office.</span></span> <span data-ttu-id="413ff-109">Если надстройка содержит собственный язык компонентов, не нужно отказываться от него в пользу Fabric.</span><span class="sxs-lookup"><span data-stu-id="413ff-109">If your add-in has its own preexisting component language, you don't need to discard it in favor of Fabric.</span></span> <span data-ttu-id="413ff-110">Найдите возможности сохранить его, интегрируя надстройку с Office.</span><span class="sxs-lookup"><span data-stu-id="413ff-110">Look for opportunities to retain it while integrating with Office.</span></span> <span data-ttu-id="413ff-111">Рассмотрите способы изменения стилистических элементов и удаления конфликтов или примените понятные пользователям стили и поведение.</span><span class="sxs-lookup"><span data-stu-id="413ff-111">Consider ways to swap out stylistic elements, remove conflicts, or adopt styles and behaviors that remove user confusion.</span></span>

<span data-ttu-id="413ff-112">Предоставленные шаблоны — это наилучшие решения, основанные на общих сценариях клиентов и исследованиях работы пользователей.</span><span class="sxs-lookup"><span data-stu-id="413ff-112">The provided patterns are best practice solutions based on common customer scenarios and user experience research.</span></span> <span data-ttu-id="413ff-113">Они призваны обеспечить как быструю отправную точку для проектирования и разработки надстроек, так и руководство для достижения баланса между элементами Майкрософт и фирменной символикой.</span><span class="sxs-lookup"><span data-stu-id="413ff-113">They are meant to provide both a quick entry point to designing and developing add-ins as well as guidance to achieve balance between Microsoft and brand elements.</span></span> <span data-ttu-id="413ff-114">Предоставление удобного и современного пользовательского интерфейса, который гармонично сочетает элементы оформления из языка дизайна Microsoft Fabric и уникальную фирменную символику партнера, может помочь лучше удерживать пользовательскую аудиторию и внедрять вашу надстройку.</span><span class="sxs-lookup"><span data-stu-id="413ff-114">Providing a clean, modern user experience that balances design elements from Microsoft's Fabric design language and the partner's unique brand identity may help increase user retention and adoption of your add-in.</span></span>

<span data-ttu-id="413ff-115">Используйте шаблонные заготовки пользовательского интерфейса для того, чтобы:</span><span class="sxs-lookup"><span data-stu-id="413ff-115">Use the UX pattern templates to:</span></span>

* <span data-ttu-id="413ff-116">применять решения в распространенных клиентских сценариях;</span><span class="sxs-lookup"><span data-stu-id="413ff-116">Apply solutions to common customer scenarios.</span></span>
* <span data-ttu-id="413ff-117">следовать рекомендациям по оформлению;</span><span class="sxs-lookup"><span data-stu-id="413ff-117">Apply design best practices.</span></span>
* <span data-ttu-id="413ff-118">внедрять компоненты и стили [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started);</span><span class="sxs-lookup"><span data-stu-id="413ff-118">Incorporate [Office UI Fabric](https://developer.microsoft.com/fabric#/get-started) components and styles.</span></span>
* <span data-ttu-id="413ff-119">создавать надстройки, внешний вид которых согласован со стандартным пользовательским интерфейсом Office;</span><span class="sxs-lookup"><span data-stu-id="413ff-119">Build add-ins that visually integrate with the default Office UI.</span></span>
* <span data-ttu-id="413ff-120">формировать и визуализировать пользовательский интерфейс.</span><span class="sxs-lookup"><span data-stu-id="413ff-120">Ideate and visualize UX.</span></span>

## <a name="getting-started"></a><span data-ttu-id="413ff-121">Начало работы</span><span class="sxs-lookup"><span data-stu-id="413ff-121">Getting started</span></span>

<span data-ttu-id="413ff-122">Шаблоны организованы по ключевым действиям или функциональным возможностям, которые часто используются в надстройке.</span><span class="sxs-lookup"><span data-stu-id="413ff-122">The patterns are organized by key actions or experiences that are common in an add-in.</span></span> <span data-ttu-id="413ff-123">Основные группы:</span><span class="sxs-lookup"><span data-stu-id="413ff-123">The main groups are:</span></span>

* [<span data-ttu-id="413ff-124">Первый запуск (FRE)</span><span class="sxs-lookup"><span data-stu-id="413ff-124">First run experience (FRE)</span></span>](../design/first-run-experience-patterns.md)
* [<span data-ttu-id="413ff-125">Проверка подлинности</span><span class="sxs-lookup"><span data-stu-id="413ff-125">Authentication</span></span>](../design/authentication-patterns.md)
* [<span data-ttu-id="413ff-126">Навигация</span><span class="sxs-lookup"><span data-stu-id="413ff-126">Navigation</span></span>](../design/navigation-patterns.md)
* [<span data-ttu-id="413ff-127">Разработка фирменной символики</span><span class="sxs-lookup"><span data-stu-id="413ff-127">Branding Design</span></span>](../design/branding-patterns.md)

<span data-ttu-id="413ff-128">Просмотрите каждую группу, чтобы получить представление о том, как можно создавать свои надстройки с использованием рекомендаций.</span><span class="sxs-lookup"><span data-stu-id="413ff-128">Browse each grouping to get an idea of how you can design your add-in using best practices.</span></span>

> [!NOTE]
> <span data-ttu-id="413ff-129">Примеры экранов, демонстрируемые в этой документации, созданы и отображены с разрешением **1366x768**.</span><span class="sxs-lookup"><span data-stu-id="413ff-129">The example screens shown throughout this documentation are designed and displayed at a resolution of **1366x768**.</span></span>

## <a name="see-also"></a><span data-ttu-id="413ff-130">См. также</span><span class="sxs-lookup"><span data-stu-id="413ff-130">See also</span></span>

* [<span data-ttu-id="413ff-131">Наборы средств оформления</span><span class="sxs-lookup"><span data-stu-id="413ff-131">Design toolkits</span></span>](design-toolkits.md)
* [<span data-ttu-id="413ff-132">Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="413ff-132">Office UI Fabric</span></span>](https://developer.microsoft.com/fabric)
* [<span data-ttu-id="413ff-133">Рекомендации по разработке надстроек Office</span><span class="sxs-lookup"><span data-stu-id="413ff-133">Best practices for developing Office Add-ins</span></span>](/office/dev/add-ins/concepts/add-in-development-best-practices)
* [<span data-ttu-id="413ff-134">Начало работы с Fabric React</span><span class="sxs-lookup"><span data-stu-id="413ff-134">Get started using Fabric React</span></span>](/office/dev/add-ins/design/using-office-ui-fabric-react)
