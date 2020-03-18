---
title: Язык оформления надстроек Office
description: Сведения о том, как сделать надстройку Office визуально совместимой с Office.
ms.date: 12/04/2017
localization_priority: Normal
ms.openlocfilehash: 0a1d175401ebaabe9c17cae18d196bc6461ba57c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718596"
---
# <a name="office-add-in-design-language"></a><span data-ttu-id="313df-103">Язык оформления надстроек Office</span><span class="sxs-lookup"><span data-stu-id="313df-103">Office Add-in design language</span></span>

<span data-ttu-id="313df-p101">Язык дизайна Office — это простая визуальная система, которая обеспечивает согласованность всех настроек. Она содержит набор визуальных элементов, определяющих интерфейсы Office, в том числе:</span><span class="sxs-lookup"><span data-stu-id="313df-p101">The Office design language is a clean and simple visual system that ensures consistency across experiences. It contains a set of visual elements that define Office interfaces, including:</span></span>

- <span data-ttu-id="313df-106">стандартный шрифт;</span><span class="sxs-lookup"><span data-stu-id="313df-106">A standard typeface</span></span>
- <span data-ttu-id="313df-107">общая цветовая палитра;</span><span class="sxs-lookup"><span data-stu-id="313df-107">A common color palette</span></span>
- <span data-ttu-id="313df-108">набор типографских размеров и весов;</span><span class="sxs-lookup"><span data-stu-id="313df-108">A set of typographic sizes and weights</span></span>
- <span data-ttu-id="313df-109">рекомендации по созданию значков;</span><span class="sxs-lookup"><span data-stu-id="313df-109">Icon guidelines</span></span>
- <span data-ttu-id="313df-110">общие ресурсы значков;</span><span class="sxs-lookup"><span data-stu-id="313df-110">Shared icon assets</span></span>
- <span data-ttu-id="313df-111">определения анимации;</span><span class="sxs-lookup"><span data-stu-id="313df-111">Animation definitions</span></span>
- <span data-ttu-id="313df-112">общие компоненты.</span><span class="sxs-lookup"><span data-stu-id="313df-112">Common components</span></span>

<span data-ttu-id="313df-p102">[Office UI Fabric](https://developer.microsoft.com/fabric) — это официальная клиентская платформа для разработки с использованием языка дизайна Office. Использовать платформу Fabric необязательно, но это самый быстрый способ обеспечить полную совместимость надстроек с Office. Воспользуйтесь преимуществами платформы Fabric для проектирования и создания надстроек, расширяющих возможности Office.</span><span class="sxs-lookup"><span data-stu-id="313df-p102">[Office UI Fabric](https://developer.microsoft.com/fabric) is the official front-end framework for building with the Office design language. Using Fabric is optional, but it is the fastest way to ensure that your add-ins feel like a natural extension of Office. Take advantage of Fabric to design and build add-ins that complement Office.</span></span>

<span data-ttu-id="313df-p103">Многие надстройки Office связаны с имеющейся фирменной символикой. В надстройке можно сохранить фирменную символику вместе с визуальным языком или языком компонентов. Найдите возможности сохранить собственный визуальный язык, интегрируя надстройку с Office. Рассмотрите возможности изменить цвета, оформление, значки или другие стилистические элементы Office на элементы собственной торговой марки. Рассмотрите способы использования распространенных макетов надстроек или конструктивных шаблонов при вставке элементов управления и компонентов, хорошо знакомых для клиентов.</span><span class="sxs-lookup"><span data-stu-id="313df-p103">Many Office Add-ins are associated with a preexisting brand. You can retain a strong brand and its visual or component language in your add-in. Look for opportunities to retain your own visual language while integrating with Office. Consider ways to swap out Office colors, typography, icons, or other stylistic elements with elements of your own brand. Consider ways to follow common add-in layouts or UX design patterns while inserting controls and components that are familiar to your customers.</span></span>

<span data-ttu-id="313df-p104">Вставка фирменного пользовательского интерфейса на основе HTML в пределах системы Office может создавать неудобства для клиентов. Найдите баланс между символикой Office и фирменной символикой вашей компании. Надстройка зачастую не вписывается в Office из-за конфликта между стилистическими элементами. Например, оформление превышает допустимый размер и выходит за пределы сетки, используемые цвета контрастируют или создают сильный шум, анимация избыточна, а ее поведение не соответствует поведению Office. Внешний вид и поведение элементов управления или компонентов значительно отличаются от стандартов Office.</span><span class="sxs-lookup"><span data-stu-id="313df-p104">Inserting a heavily branded HTML-based UI inside of Office can create dissonance for customers. Find a balance that fits seamlessly in Office but also clearly aligns with your service or parent brand. When an add-in does not fit with Office, it's often because stylistic elements conflict. For example, typography is too large and off grid, colors are contrasting or particularly loud, or animations are superfluous and behave differently than Office. The appearance and behavior of controls or components veer too far from Office standards.</span></span>
