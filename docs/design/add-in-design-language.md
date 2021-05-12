---
title: Язык оформления надстроек Office
description: Узнайте, как сделать Office надстройку визуально совместимой с Office.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 4374eaad1e660728a347d19a012d345b642755af
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330061"
---
# <a name="office-add-in-design-language"></a><span data-ttu-id="24940-103">Язык оформления надстроек Office</span><span class="sxs-lookup"><span data-stu-id="24940-103">Office Add-in design language</span></span>

<span data-ttu-id="24940-p101">Язык дизайна Office — это простая визуальная система, которая обеспечивает согласованность всех настроек. Она содержит набор визуальных элементов, определяющих интерфейсы Office, в том числе:</span><span class="sxs-lookup"><span data-stu-id="24940-p101">The Office design language is a clean and simple visual system that ensures consistency across experiences. It contains a set of visual elements that define Office interfaces, including:</span></span>

- <span data-ttu-id="24940-106">стандартный шрифт;</span><span class="sxs-lookup"><span data-stu-id="24940-106">A standard typeface</span></span>
- <span data-ttu-id="24940-107">общая цветовая палитра;</span><span class="sxs-lookup"><span data-stu-id="24940-107">A common color palette</span></span>
- <span data-ttu-id="24940-108">набор типографских размеров и весов;</span><span class="sxs-lookup"><span data-stu-id="24940-108">A set of typographic sizes and weights</span></span>
- <span data-ttu-id="24940-109">рекомендации по созданию значков;</span><span class="sxs-lookup"><span data-stu-id="24940-109">Icon guidelines</span></span>
- <span data-ttu-id="24940-110">общие ресурсы значков;</span><span class="sxs-lookup"><span data-stu-id="24940-110">Shared icon assets</span></span>
- <span data-ttu-id="24940-111">определения анимации;</span><span class="sxs-lookup"><span data-stu-id="24940-111">Animation definitions</span></span>
- <span data-ttu-id="24940-112">общие компоненты.</span><span class="sxs-lookup"><span data-stu-id="24940-112">Common components</span></span>

<span data-ttu-id="24940-113">[Fluent UI](../design/add-in-design.md) — это официальная интерфейсная основа для создания с Office языка разработки.</span><span class="sxs-lookup"><span data-stu-id="24940-113">[Fluent UI](../design/add-in-design.md) is the official front-end framework for building with the Office design language.</span></span> <span data-ttu-id="24940-114">Использование пользовательского интерфейса Fluent является необязательным, но это самый быстрый способ обеспечить, чтобы ваши надстройки чувствовали себя естественным расширением Office.</span><span class="sxs-lookup"><span data-stu-id="24940-114">Using Fluent UI is optional, but it is the fastest way to ensure that your add-ins feel like a natural extension of Office.</span></span> <span data-ttu-id="24940-115">Создайте и создайте надстройки, которые дополняют Office.</span><span class="sxs-lookup"><span data-stu-id="24940-115">Take advantage of Fluent UI to design and build add-ins that complement Office.</span></span>

<span data-ttu-id="24940-p103">Многие надстройки Office связаны с имеющейся фирменной символикой. В надстройке можно сохранить фирменную символику вместе с визуальным языком или языком компонентов. Найдите возможности сохранить собственный визуальный язык, интегрируя надстройку с Office. Рассмотрите возможности изменить цвета, оформление, значки или другие стилистические элементы Office на элементы собственной торговой марки. Рассмотрите способы использования распространенных макетов надстроек или конструктивных шаблонов при вставке элементов управления и компонентов, хорошо знакомых для клиентов.</span><span class="sxs-lookup"><span data-stu-id="24940-p103">Many Office Add-ins are associated with a preexisting brand. You can retain a strong brand and its visual or component language in your add-in. Look for opportunities to retain your own visual language while integrating with Office. Consider ways to swap out Office colors, typography, icons, or other stylistic elements with elements of your own brand. Consider ways to follow common add-in layouts or UX design patterns while inserting controls and components that are familiar to your customers.</span></span>

<span data-ttu-id="24940-p104">Вставка фирменного пользовательского интерфейса на основе HTML в пределах системы Office может создавать неудобства для клиентов. Найдите баланс между символикой Office и фирменной символикой вашей компании. Надстройка зачастую не вписывается в Office из-за конфликта между стилистическими элементами. Например, оформление превышает допустимый размер и выходит за пределы сетки, используемые цвета контрастируют или создают сильный шум, анимация избыточна, а ее поведение не соответствует поведению Office. Внешний вид и поведение элементов управления или компонентов значительно отличаются от стандартов Office.</span><span class="sxs-lookup"><span data-stu-id="24940-p104">Inserting a heavily branded HTML-based UI inside of Office can create dissonance for customers. Find a balance that fits seamlessly in Office but also clearly aligns with your service or parent brand. When an add-in does not fit with Office, it's often because stylistic elements conflict. For example, typography is too large and off grid, colors are contrasting or particularly loud, or animations are superfluous and behave differently than Office. The appearance and behavior of controls or components veer too far from Office standards.</span></span>
