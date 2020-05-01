---
title: Рекомендации по использованию значка стиля "inline" для надстроек Office
description: Ознакомьтесь с рекомендациями по использованию значков нелинейного стиля в надстройках Office.
ms.date: 12/09/2019
localization_priority: Normal
ms.openlocfilehash: 8d1ac2dff76b852cd69b38bd2c138d1ba43f203c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718603"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a><span data-ttu-id="90d39-103">Рекомендации по использованию значка стиля "inline" для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="90d39-103">Monoline style icon guidelines for Office Add-ins</span></span>

<span data-ttu-id="90d39-104">Стиль значки, который используется в Office 365.</span><span class="sxs-lookup"><span data-stu-id="90d39-104">Monoline style iconography are used in Office 365.</span></span> <span data-ttu-id="90d39-105">Если вы предпочитаете, чтобы значки выглядели как неактуальный стиль Office 2013, не относящегося к подписке, обратитесь к разделу [новые рекомендации по использованию значков стилей для надстроек Office](add-in-icons-fresh.md).</span><span class="sxs-lookup"><span data-stu-id="90d39-105">If you would prefer that your icons match the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

## <a name="office-monoline-visual-style"></a><span data-ttu-id="90d39-106">Линейный визуальный стиль Office</span><span class="sxs-lookup"><span data-stu-id="90d39-106">Office Monoline visual style</span></span>

<span data-ttu-id="90d39-107">Цель стиля "линейный" для обеспечения согласованных, ясных и доступных значки для общения действий и функций с простыми визуальными элементами, обеспечения доступности значков для всех пользователей и стиля, согласованного с теми, которые используются в других окнах Windows.</span><span class="sxs-lookup"><span data-stu-id="90d39-107">The goal of the Monoline style to have consistent, clear, and accessible iconography to communicate action and features with simple visuals, ensure the icons are accessible to all users, and have a style that is consistent with those used elsewhere in Windows.</span></span>

<span data-ttu-id="90d39-108">Следующие рекомендации предназначены для сторонних разработчиков, которые хотят создать значки для функций, которые будут согласованы с уже присутствующими продуктами Office.</span><span class="sxs-lookup"><span data-stu-id="90d39-108">The following guidelines are for 3rd party developers who want to create icons for features that will be consistent with the icons already present Office products.</span></span>

### <a name="design-principles"></a><span data-ttu-id="90d39-109">Принципы разработки</span><span class="sxs-lookup"><span data-stu-id="90d39-109">Design principles</span></span>

-   <span data-ttu-id="90d39-110">Простой, чистый, ясный.</span><span class="sxs-lookup"><span data-stu-id="90d39-110">Simple, clean, clear.</span></span>
-   <span data-ttu-id="90d39-111">Содержать только необходимые элементы.</span><span class="sxs-lookup"><span data-stu-id="90d39-111">Contain only necessary elements.</span></span>
-   <span data-ttu-id="90d39-112">Стиль значков Windows.</span><span class="sxs-lookup"><span data-stu-id="90d39-112">Inspired by Windows icon style.</span></span>
-   <span data-ttu-id="90d39-113">Доступен всем пользователям.</span><span class="sxs-lookup"><span data-stu-id="90d39-113">Accessible to all users.</span></span>

#### <a name="conveying-meaning"></a><span data-ttu-id="90d39-114">Передающееся значение</span><span class="sxs-lookup"><span data-stu-id="90d39-114">Conveying meaning</span></span>

-   <span data-ttu-id="90d39-115">Используйте элементы с описанием, например страницу, чтобы представить документ или конверт для представления почты.</span><span class="sxs-lookup"><span data-stu-id="90d39-115">Use descriptive elements such as a page to represent a document or an envelope to represent mail.</span></span>
-   <span data-ttu-id="90d39-116">Используйте один и тот же элемент для представления той же концепции, т.е. почта всегда представлена конвертом, а не штампом.</span><span class="sxs-lookup"><span data-stu-id="90d39-116">Use the same element to represent the same concept, i.e., mail is always represented by an envelope, not a stamp.</span></span>
-   <span data-ttu-id="90d39-117">Используйте базовую метафору во время разработки концепции.</span><span class="sxs-lookup"><span data-stu-id="90d39-117">Use a core metaphor during concept development.</span></span>

#### <a name="reduction-of-elements"></a><span data-ttu-id="90d39-118">Сокращение элементов</span><span class="sxs-lookup"><span data-stu-id="90d39-118">Reduction of Elements</span></span>

-   <span data-ttu-id="90d39-119">Сократите значок до основного значения, используя только те элементы, которые необходимы для метафоры.</span><span class="sxs-lookup"><span data-stu-id="90d39-119">Reduce the icon to its core meaning, using only elements that are essential to the metaphor.</span></span>
-   <span data-ttu-id="90d39-120">Ограничьте количество элементов в значке двумя, независимо от размера значка.</span><span class="sxs-lookup"><span data-stu-id="90d39-120">Limit the number of elements in an icon to two, regardless of icon size.</span></span>

#### <a name="consistency"></a><span data-ttu-id="90d39-121">Обнаружен</span><span class="sxs-lookup"><span data-stu-id="90d39-121">Consistency</span></span>

<span data-ttu-id="90d39-122">Размеры, расположение и цвет значков должны быть согласованы.</span><span class="sxs-lookup"><span data-stu-id="90d39-122">Sizes, arrangement, and color of icons should be consistent.</span></span>

#### <a name="styling"></a><span data-ttu-id="90d39-123">Изменении</span><span class="sxs-lookup"><span data-stu-id="90d39-123">Styling</span></span>

##### <a name="perspective"></a><span data-ttu-id="90d39-124">Perspective</span><span class="sxs-lookup"><span data-stu-id="90d39-124">Perspective</span></span>

<span data-ttu-id="90d39-125">По умолчанию значки с фиксированной линейкой перемещаются вперед.</span><span class="sxs-lookup"><span data-stu-id="90d39-125">Monoline icons are forward-facing by default.</span></span> <span data-ttu-id="90d39-126">Некоторые элементы, требующие перспективы и/или вращения, такие как куб, разрешены, но исключения должны быть сохранены как минимум.</span><span class="sxs-lookup"><span data-stu-id="90d39-126">Certain elements that require perspective and/or rotation, such as a cube, are allowed, but exceptions should be kept to a minimum.</span></span>

##### <a name="embellishment"></a><span data-ttu-id="90d39-127">Надстрочные знаки</span><span class="sxs-lookup"><span data-stu-id="90d39-127">Embellishment</span></span>

<span data-ttu-id="90d39-128">"Однострочный" — чистый простой стиль.</span><span class="sxs-lookup"><span data-stu-id="90d39-128">Monoline is a clean minimal style.</span></span> <span data-ttu-id="90d39-129">Все использует плоский цвет, что означает, что нет градиентов, текстур или источников света.</span><span class="sxs-lookup"><span data-stu-id="90d39-129">Everything uses flat color, which means there are no gradients, textures, or light sources.</span></span>

## <a name="designing"></a><span data-ttu-id="90d39-130">Работ</span><span class="sxs-lookup"><span data-stu-id="90d39-130">Designing</span></span>

### <a name="sizes"></a><span data-ttu-id="90d39-131">Масштаба</span><span class="sxs-lookup"><span data-stu-id="90d39-131">Sizes</span></span>

<span data-ttu-id="90d39-132">Для поддержки устройств с высоким разрешением рекомендуется создать каждый значок на всех этих размерах.</span><span class="sxs-lookup"><span data-stu-id="90d39-132">We recommend that you produce each icon in all these sizes to support high DPI devices.</span></span> <span data-ttu-id="90d39-133">Крайне *обязательные* размеры — 16px, 20px и интервалами по 32, так как размер 100%.</span><span class="sxs-lookup"><span data-stu-id="90d39-133">The absolutely *required* sizes are 16px, 20px, and 32px, as those are the 100% sizes.</span></span>

<span data-ttu-id="90d39-134">**16px, 20px, интервалами по 24, интервалами по 32, 40px, 48px, 64px, 80px, 96px**</span><span class="sxs-lookup"><span data-stu-id="90d39-134">**16px, 20px, 24px, 32px, 40px, 48px, 64px, 80px, 96px**</span></span>

### <a name="layout"></a><span data-ttu-id="90d39-135">Макет</span><span class="sxs-lookup"><span data-stu-id="90d39-135">Layout</span></span>

<span data-ttu-id="90d39-136">Ниже приведен пример макета значков с модификатором.</span><span class="sxs-lookup"><span data-stu-id="90d39-136">The following is an example of icon layout with a modifier.</span></span>

![Пример значка с модификатором](../images/monolineicon1.png)  ![Тот же пример, в котором есть фоновые выноски сетки для базового, модификатора, заполнения и отреза.](../images/monolineicon2.png)

#### <a name="elements"></a><span data-ttu-id="90d39-139">Элементы</span><span class="sxs-lookup"><span data-stu-id="90d39-139">Elements</span></span>

- <span data-ttu-id="90d39-140">**Основание**: основная концепция, которую представляет значок.</span><span class="sxs-lookup"><span data-stu-id="90d39-140">**Base**: The main concept that the icon represents.</span></span> <span data-ttu-id="90d39-141">Обычно это единственный визуальный элемент, который требуется для значка, но иногда его можно улучшить с помощью дополнительного элемента, модификатора.</span><span class="sxs-lookup"><span data-stu-id="90d39-141">This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.</span></span>

- <span data-ttu-id="90d39-142">**Модификатор** Любой элемент, перекрывающих базовый; то есть модификатор, который обычно представляет действие или состояние.</span><span class="sxs-lookup"><span data-stu-id="90d39-142">**Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status.</span></span> <span data-ttu-id="90d39-143">Он изменяет базовый элемент, выполняя в качестве дополнения, изменения или дескриптора.</span><span class="sxs-lookup"><span data-stu-id="90d39-143">It modifies the base element by acting as an addition, alteration, or a descriptor.</span></span>

![Сетка с областями базовой области и модификаторов.](../images/monolineicon3.png)

### <a name="construction"></a><span data-ttu-id="90d39-145">Строительство</span><span class="sxs-lookup"><span data-stu-id="90d39-145">Construction</span></span>

#### <a name="element-placement"></a><span data-ttu-id="90d39-146">Размещение элементов</span><span class="sxs-lookup"><span data-stu-id="90d39-146">Element placement</span></span>

<span data-ttu-id="90d39-147">Базовые элементы размещаются в центре значка в пределах заполнения.</span><span class="sxs-lookup"><span data-stu-id="90d39-147">Base elements are placed in the center of the icon within the padding.</span></span> <span data-ttu-id="90d39-148">Если он не может быть разработано по центру, то основной правый раздел должен находиться в начале.</span><span class="sxs-lookup"><span data-stu-id="90d39-148">If it can't be placed perfectly centered, then the base should err to the top right.</span></span> <span data-ttu-id="90d39-149">В следующем примере значок идеально выравнивается по центру:</span><span class="sxs-lookup"><span data-stu-id="90d39-149">In the following example, the icon is perfectly centered:</span></span>

![Изображение с точно выровненным по центру значком](../images/monolineicon4.png)

<span data-ttu-id="90d39-151">В следующем примере значок ерринг слева.</span><span class="sxs-lookup"><span data-stu-id="90d39-151">In the following example, the icon is erring to the left.</span></span>

![Изображение значка, еррс влево](../images/monolineicon5.png)

<span data-ttu-id="90d39-153">Модификаторы почти всегда располагаются в правом нижнем углу холста значка.</span><span class="sxs-lookup"><span data-stu-id="90d39-153">Modifiers are almost always placed in the bottom right corner of the icon canvas.</span></span> <span data-ttu-id="90d39-154">В некоторых редких случаях модификаторы размещаются в другой угол.</span><span class="sxs-lookup"><span data-stu-id="90d39-154">In some rare cases, modifiers are placed in a different corner.</span></span> <span data-ttu-id="90d39-155">Например, если базовый элемент не распознается с помощью модификатора в правом нижнем углу, его можно разместить в левом верхнем углу.</span><span class="sxs-lookup"><span data-stu-id="90d39-155">For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.</span></span>

![Изображение, на котором показаны несколько значков с модификатором в нижнем правом углу, но с модификатором в верхнем левом углу](../images/monolineicon6.png)

#### <a name="padding"></a><span data-ttu-id="90d39-157">Внутренние поля</span><span class="sxs-lookup"><span data-stu-id="90d39-157">Padding</span></span>

<span data-ttu-id="90d39-158">Каждый значок размера имеет заданный объем заполнения вокруг значка.</span><span class="sxs-lookup"><span data-stu-id="90d39-158">Each size icon has a specified amount of padding around the icon.</span></span> <span data-ttu-id="90d39-159">Базовый элемент остается в пределах заполнения, но модификатор должен Бутт до края холста, расширяя за пределы заполнения---до края границы значка.</span><span class="sxs-lookup"><span data-stu-id="90d39-159">The base element stays within the padding, but the modifier should butt up to the edge of the canvas, extending outside of the padding---to the edge of the icon border.</span></span> <span data-ttu-id="90d39-160">На следующих изображениях показана Рекомендуемая величина заполнения, используемая для каждого размера значков.</span><span class="sxs-lookup"><span data-stu-id="90d39-160">The following images show the recommended padding to use for each of the icon sizes.</span></span>

|<span data-ttu-id="90d39-161">**16 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-161">**16px**</span></span>|<span data-ttu-id="90d39-162">**20 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-162">**20px**</span></span>|<span data-ttu-id="90d39-163">**24 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-163">**24px**</span></span>|<span data-ttu-id="90d39-164">**32 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-164">**32px**</span></span>|<span data-ttu-id="90d39-165">**40 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-165">**40px**</span></span>|<span data-ttu-id="90d39-166">**48 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-166">**48px**</span></span>|<span data-ttu-id="90d39-167">**64 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-167">**64px**</span></span>|<span data-ttu-id="90d39-168">**80 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-168">**80px**</span></span>|<span data-ttu-id="90d39-169">**96px**</span><span class="sxs-lookup"><span data-stu-id="90d39-169">**96px**</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![16 точек](../images/monolineicon7.png)|![значок 20 точек](../images/monolineicon8.png)|![значок 24 ПКС](../images/monolineicon9.png)|![32 точек](../images/monolineicon10.png)|![40 точек](../images/monolineicon11.png)|![48 точек](../images/monolineicon12.png)|![64 точек](../images/monolineicon13.png)|![80 точек](../images/monolineicon14.png)|![96 точек](../images/monolineicon15.png)|

#### <a name="line-weights"></a><span data-ttu-id="90d39-179">Толщина линий</span><span class="sxs-lookup"><span data-stu-id="90d39-179">Line weights</span></span>

<span data-ttu-id="90d39-180">"Inline" — это стиль, который облагаются строкой и контурными фигурами.</span><span class="sxs-lookup"><span data-stu-id="90d39-180">Monoline is a style dominated by line and outlined shapes.</span></span> <span data-ttu-id="90d39-181">В зависимости от размера, который вы создаете значок, должен использовать следующие веса линии.</span><span class="sxs-lookup"><span data-stu-id="90d39-181">Depending on what size you are producing the icon should use the following line weights.</span></span>

|<span data-ttu-id="90d39-182">**Размер значка:**</span><span class="sxs-lookup"><span data-stu-id="90d39-182">**Icon Size:**</span></span>|<span data-ttu-id="90d39-183">**16 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-183">**16px**</span></span>|<span data-ttu-id="90d39-184">**20 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-184">**20px**</span></span>|<span data-ttu-id="90d39-185">**24 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-185">**24px**</span></span>|<span data-ttu-id="90d39-186">**32 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-186">**32px**</span></span>|<span data-ttu-id="90d39-187">**40 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-187">**40px**</span></span>|<span data-ttu-id="90d39-188">**48 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-188">**48px**</span></span>|<span data-ttu-id="90d39-189">**64 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-189">**64px**</span></span>|<span data-ttu-id="90d39-190">**80 пк**</span><span class="sxs-lookup"><span data-stu-id="90d39-190">**80px**</span></span>|<span data-ttu-id="90d39-191">**96px**</span><span class="sxs-lookup"><span data-stu-id="90d39-191">**96px**</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|<span data-ttu-id="90d39-192">**Толщина линии:**</span><span class="sxs-lookup"><span data-stu-id="90d39-192">**Line Weight:**</span></span>|<span data-ttu-id="90d39-193">1 пк</span><span class="sxs-lookup"><span data-stu-id="90d39-193">1px</span></span>|<span data-ttu-id="90d39-194">1 пк</span><span class="sxs-lookup"><span data-stu-id="90d39-194">1px</span></span>|<span data-ttu-id="90d39-195">1 пк</span><span class="sxs-lookup"><span data-stu-id="90d39-195">1px</span></span>|<span data-ttu-id="90d39-196">1 пк</span><span class="sxs-lookup"><span data-stu-id="90d39-196">1px</span></span>|<span data-ttu-id="90d39-197">2 пк</span><span class="sxs-lookup"><span data-stu-id="90d39-197">2px</span></span>|<span data-ttu-id="90d39-198">2 пк</span><span class="sxs-lookup"><span data-stu-id="90d39-198">2px</span></span>|<span data-ttu-id="90d39-199">2 пк</span><span class="sxs-lookup"><span data-stu-id="90d39-199">2px</span></span>|<span data-ttu-id="90d39-200">2 пк</span><span class="sxs-lookup"><span data-stu-id="90d39-200">2px</span></span>|<span data-ttu-id="90d39-201">3 пк</span><span class="sxs-lookup"><span data-stu-id="90d39-201">3px</span></span>|
||![16 точек](../images/monolineicon16.png)|![значок 20 точек](../images/monolineicon17.png)|![значок 24 ПКС](../images/monolineicon18.png)|![32 точек](../images/monolineicon19.png)|![40 точек](../images/monolineicon20.png)|![48 точек](../images/monolineicon21.png)|![64 точек](../images/monolineicon22.png)|![80 точек](../images/monolineicon23.png)|![96 точек](../images/monolineicon24.png)|

#### <a name="cutouts"></a><span data-ttu-id="90d39-211">Контуры</span><span class="sxs-lookup"><span data-stu-id="90d39-211">Cutouts</span></span>

<span data-ttu-id="90d39-212">Когда элемент Icon помещается поверх другого элемента, используется отрезки (элемента нижнего элемента) для предоставления промежутка между двумя элементами, в основном для удобства чтения.</span><span class="sxs-lookup"><span data-stu-id="90d39-212">When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes.</span></span> <span data-ttu-id="90d39-213">Обычно это происходит, когда модификатор помещается поверх базового элемента, но существуют также случаи, когда ни один из элементов не является модификатором.</span><span class="sxs-lookup"><span data-stu-id="90d39-213">This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier.</span></span> <span data-ttu-id="90d39-214">Эти отрезки между двумя элементами иногда называют "пропуском".</span><span class="sxs-lookup"><span data-stu-id="90d39-214">These cutouts between the two elements is sometimes referred to as a "gap".</span></span>

<span data-ttu-id="90d39-215">Размер зазора должен совпадать с шириной линии, используемой для этого размера.</span><span class="sxs-lookup"><span data-stu-id="90d39-215">The size of the gap should be the same width as the line weight used on that size.</span></span> <span data-ttu-id="90d39-216">При создании значка 16px ширина зазора будет 1 ПКС, а если это значок 48px, то зазор должен быть 2 ПКС.</span><span class="sxs-lookup"><span data-stu-id="90d39-216">If making a 16px icon, the gap width would be 1px and if it is a 48px icon then the gap should be 2px.</span></span> <span data-ttu-id="90d39-217">В следующем примере показан значок интервалами по 32 с разрывом 1 ПКС между модификатором и базовым основанием.</span><span class="sxs-lookup"><span data-stu-id="90d39-217">The following example shows a 32px icon with a gap of 1px between the modifier and the underlying base.</span></span>

![значок интервалами по 32 с пропуском 1 ПКС между модификатором и базовым базовым](../images/monolineicon25.png)

<span data-ttu-id="90d39-219">В некоторых случаях зазор может быть увеличен на 1/2 ПКС, если у модификатора есть диагональный или изогнутый край, а стандартный зазор не обеспечивает достаточного расстояния.</span><span class="sxs-lookup"><span data-stu-id="90d39-219">In some cases, the gap can be increase by a 1/2px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation.</span></span> <span data-ttu-id="90d39-220">Скорее всего, они повлияют только на значки с 1 пксной толщиной линии; 16px, 20px, интервалами по 24 и интервалами по 32.</span><span class="sxs-lookup"><span data-stu-id="90d39-220">This will likely only affect the icons with 1px line weight; 16px, 20px, 24px, and 32px.</span></span>

#### <a name="background-fills"></a><span data-ttu-id="90d39-221">Заливка фона</span><span class="sxs-lookup"><span data-stu-id="90d39-221">Background fills</span></span>

<span data-ttu-id="90d39-222">Для большинства значков в наборе значков в виде линии требуются фоновые заливки.</span><span class="sxs-lookup"><span data-stu-id="90d39-222">Most icons in the Monoline icon set require background fills.</span></span> <span data-ttu-id="90d39-223">Однако в некоторых случаях нет необходимости применять заливку для объекта.</span><span class="sxs-lookup"><span data-stu-id="90d39-223">However, there are cases where the object would not naturally have a fill, so no fill should be applied.</span></span> <span data-ttu-id="90d39-224">Следующие значки имеют белую заливку:</span><span class="sxs-lookup"><span data-stu-id="90d39-224">The following icons have a white fill:</span></span>

![Пять значков имеют белую заливку](../images/monolineicon26.png)

<span data-ttu-id="90d39-226">Следующие значки не имеют заливки.</span><span class="sxs-lookup"><span data-stu-id="90d39-226">The following icons have no fill.</span></span> <span data-ttu-id="90d39-227">(Значок шестеренки включается, чтобы показать, что не заполнено Центральная дыра.) ![Пять значков без заливки](../images/monolineicon27.png)</span><span class="sxs-lookup"><span data-stu-id="90d39-227">(The gear icon is included to show that the center hole is not filled.) ![Five icons with no fill](../images/monolineicon27.png)</span></span>

##### <a name="best-practices-for-fills"></a><span data-ttu-id="90d39-228">Рекомендации по заполнению</span><span class="sxs-lookup"><span data-stu-id="90d39-228">Best practices for fills</span></span>

###### <a name="dos"></a><span data-ttu-id="90d39-229">Задач</span><span class="sxs-lookup"><span data-stu-id="90d39-229">Dos:</span></span>

- <span data-ttu-id="90d39-230">Заполните любой элемент, который имеет определенную границу, и, естественно, имеет заливку.</span><span class="sxs-lookup"><span data-stu-id="90d39-230">Fill any element that has a defined boundary, and would naturally have a fill.</span></span>
- <span data-ttu-id="90d39-231">Используйте отдельную фигуру, чтобы создать фоновую заливку.</span><span class="sxs-lookup"><span data-stu-id="90d39-231">Use a separate shape to create the background fill.</span></span>
- <span data-ttu-id="90d39-232">Используйте **фоновую заливку** из [цветовой палитры](#color).</span><span class="sxs-lookup"><span data-stu-id="90d39-232">Use **Background Fill** from the [color palette](#color).</span></span>
- <span data-ttu-id="90d39-233">Поддерживать разделение точек между перекрывающимися элементами.</span><span class="sxs-lookup"><span data-stu-id="90d39-233">Maintain the pixel separation between overlapping elements.</span></span>
- <span data-ttu-id="90d39-234">Заливка между несколькими объектами.</span><span class="sxs-lookup"><span data-stu-id="90d39-234">Fill between multiple objects.</span></span>

###### <a name="donts"></a><span data-ttu-id="90d39-235">Запреты</span><span class="sxs-lookup"><span data-stu-id="90d39-235">Don'ts:</span></span>

- <span data-ttu-id="90d39-236">Не заполняйте объекты, которые не должны быть заполнены. Например, скрепка.</span><span class="sxs-lookup"><span data-stu-id="90d39-236">Don't fill objects that would not naturally be filled; for example, a paperclip.</span></span>
- <span data-ttu-id="90d39-237">Не заполняйте заполнять скобки.</span><span class="sxs-lookup"><span data-stu-id="90d39-237">Don't fill brackets.</span></span>
- <span data-ttu-id="90d39-238">Не заполняйте заливку за пределами чисел или буквенных символов.</span><span class="sxs-lookup"><span data-stu-id="90d39-238">Don't fill behind numbers or alpha characters.</span></span>

### <a name="color"></a><span data-ttu-id="90d39-239">Цвет</span><span class="sxs-lookup"><span data-stu-id="90d39-239">Color</span></span>

<span data-ttu-id="90d39-240">Цветовая палитра разработана для простоты и специальных возможностей.</span><span class="sxs-lookup"><span data-stu-id="90d39-240">The color palette has been designed for simplicity and accessibility.</span></span> <span data-ttu-id="90d39-241">Он содержит 4 нейтральные цвета и два варианта для синего, зеленого, желтого, красного и фиолетового.</span><span class="sxs-lookup"><span data-stu-id="90d39-241">It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple.</span></span> <span data-ttu-id="90d39-242">Оранжевый цвет, намеренно не включен в цветовую палитру значков в виде строки.</span><span class="sxs-lookup"><span data-stu-id="90d39-242">Orange is intentionally not included in the Monoline icon color palette.</span></span> <span data-ttu-id="90d39-243">Каждый цвет предназначен для определенных способов, как описано в этом разделе.</span><span class="sxs-lookup"><span data-stu-id="90d39-243">Each color is intended to be used in specific ways as outlined in this section.</span></span>

#### <a name="palette"></a><span data-ttu-id="90d39-244">Произвольная</span><span class="sxs-lookup"><span data-stu-id="90d39-244">Palette</span></span>

![Четыре оттенка серого в виде линий](../images/monoline-grayshades.png)

![Цветовая палитра в режиме "однострочный"](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a><span data-ttu-id="90d39-247">Использование цвета</span><span class="sxs-lookup"><span data-stu-id="90d39-247">How to use color</span></span>

<span data-ttu-id="90d39-248">В цветовой палитре все цвета имеют отдельные варианты, структуры и заливки.</span><span class="sxs-lookup"><span data-stu-id="90d39-248">In the Monoline color palette, all colors have Standalone, Outline, and Fill variations.</span></span> <span data-ttu-id="90d39-249">Как правило, элементы создаются с заливкой и границей.</span><span class="sxs-lookup"><span data-stu-id="90d39-249">Generally, elements are constructed with a fill and a border.</span></span> <span data-ttu-id="90d39-250">Цвета применяются в одном из следующих шаблонов:</span><span class="sxs-lookup"><span data-stu-id="90d39-250">The colors are applied in one of the following patterns:</span></span>

- <span data-ttu-id="90d39-251">Отдельный цвет для объектов, не имеющих заливки.</span><span class="sxs-lookup"><span data-stu-id="90d39-251">The Standalone color alone for objects that have no fill.</span></span>
- <span data-ttu-id="90d39-252">Рамка использует цвет контура, а заливка использует цвет заливки.</span><span class="sxs-lookup"><span data-stu-id="90d39-252">The border uses the Outline color and the fill uses the Fill color.</span></span>
- <span data-ttu-id="90d39-253">Граница использует отдельный цвет, а заливка использует цвет заливки фона.</span><span class="sxs-lookup"><span data-stu-id="90d39-253">The border uses the Standalone color and the fill uses the Background Fill color.</span></span>

<span data-ttu-id="90d39-254">Ниже приведены примеры использования Color.</span><span class="sxs-lookup"><span data-stu-id="90d39-254">The following are examples of using color.</span></span>

![Три значка с цветом границы или заливки или и то, и другое.](../images/monolineicon28.png)

<span data-ttu-id="90d39-256">Наиболее распространенной ситуацией будет использование темно-серого элемента с заливкой фона.</span><span class="sxs-lookup"><span data-stu-id="90d39-256">The most common situation will be to have an element use Dark Gray Standalone with Background Fill.</span></span>

<span data-ttu-id="90d39-257">При использовании цветной заливки он должен всегда соответствовать соответствующему цвету контура.</span><span class="sxs-lookup"><span data-stu-id="90d39-257">When using a colored Fill, it should always be with its corresponding Outline color.</span></span> <span data-ttu-id="90d39-258">Например, синяя заливка должна использоваться только с синей структурой.</span><span class="sxs-lookup"><span data-stu-id="90d39-258">For example, Blue Fill should only be used with Blue Outline.</span></span> <span data-ttu-id="90d39-259">Но существует два исключения из этого общего правила:</span><span class="sxs-lookup"><span data-stu-id="90d39-259">But there are two exceptions to this general rule:</span></span>

- <span data-ttu-id="90d39-260">Фоновую заливку можно использовать с отдельными цветами.</span><span class="sxs-lookup"><span data-stu-id="90d39-260">Background Fill can be used with any color Standalone.</span></span>
- <span data-ttu-id="90d39-261">Светло-серая заливка можно использовать с двумя различными цветовыми контурами: темно-серый или средний серый.</span><span class="sxs-lookup"><span data-stu-id="90d39-261">Light Gray Fill can be used with two different Outline colors: Dark Gray or Medium Gray.</span></span>

#### <a name="when-to-use-color"></a><span data-ttu-id="90d39-262">Когда следует использовать Color</span><span class="sxs-lookup"><span data-stu-id="90d39-262">When to use color</span></span>

<span data-ttu-id="90d39-263">Цвет должен использоваться для передачи значения значка, а не для надстрочных знаков.</span><span class="sxs-lookup"><span data-stu-id="90d39-263">Color should be used to convey the meaning of the icon rather than for embellishment.</span></span> <span data-ttu-id="90d39-264">Он должен **выделить действие** для пользователя.</span><span class="sxs-lookup"><span data-stu-id="90d39-264">It should **highlight the action** to the user.</span></span> <span data-ttu-id="90d39-265">Когда в базовый элемент, имеющий цвет, добавляется модификатор, базовый элемент обычно включается в темно-серый и фоновую заливку, чтобы модификатор мог быть элементом Color, например, с помощью модификатора "X", добавляемого к разделу "изображение" в крайнем левом значке следующего набора.</span><span class="sxs-lookup"><span data-stu-id="90d39-265">When a modifier is added to a base element that has color, the base element is typically turned into Dark Gray and Background Fill so that the modifier can be the element of color, such as the case below with the "X" modifier being added to the picture base in the leftmost icon of the following set.</span></span>

![Пять значков, которые используют цвет](../images/monolineicon29.png)

<span data-ttu-id="90d39-267">Вы должны ограничить значки **одним** дополнительным цветом, кроме контура и закрашивания, упомянутого выше.</span><span class="sxs-lookup"><span data-stu-id="90d39-267">You should limit your icons to **one** additional color, other than the Outline and Fill mentioned above.</span></span> <span data-ttu-id="90d39-268">Однако можно использовать дополнительные цвета, если это важно для метафоры, с предельным числом двух дополнительных цветов, отличных от серого.</span><span class="sxs-lookup"><span data-stu-id="90d39-268">However, more colors can be used if it is vital for its metaphor, with a limit of two additional colors other than gray.</span></span> <span data-ttu-id="90d39-269">В редких случаях существуют исключения, когда требуется больше цветов.</span><span class="sxs-lookup"><span data-stu-id="90d39-269">In rare cases, there are exceptions when more colors are needed.</span></span> <span data-ttu-id="90d39-270">Ниже приведены хорошие примеры значков, использующих только один цвет.</span><span class="sxs-lookup"><span data-stu-id="90d39-270">The following are good examples of icons that use just one color.</span></span>

  ![Изображение из пяти значков с одним цветом](../images/monolineicon30.png)

<span data-ttu-id="90d39-272">Но следующие значки используют слишком много цветов.</span><span class="sxs-lookup"><span data-stu-id="90d39-272">But the following icons use too many colors.</span></span>

  ![Изображение из пяти значков с несколькими цветами](../images/monolineicon31.png)


<span data-ttu-id="90d39-274">Используйте **средний серый цвет** для внутреннего "содержимого", например линий сетки, в виде значка электронной таблицы.</span><span class="sxs-lookup"><span data-stu-id="90d39-274">Use **Medium Gray** for interior "content", such as grid lines in an icon of a spreadsheet.</span></span> <span data-ttu-id="90d39-275">Дополнительные внутренние цвета используются, когда контент должен показывать поведение элемента управления.</span><span class="sxs-lookup"><span data-stu-id="90d39-275">Additional interior colors are used when the content needs to show the behavior of the control.</span></span>

![Пять значков со средним серым внутренним элементами](../images/monolineicon32.png)

#### <a name="text-lines"></a><span data-ttu-id="90d39-277">Строки текста</span><span class="sxs-lookup"><span data-stu-id="90d39-277">Text lines</span></span>

<span data-ttu-id="90d39-278">Если текстовые строки находятся в контейнере (например, текст в документе), используйте средний серый цвет.</span><span class="sxs-lookup"><span data-stu-id="90d39-278">When text lines are in a "container" (for example, text on a document), use medium gray.</span></span> <span data-ttu-id="90d39-279">Текстовые строки, отсутствующие в контейнере, должны быть **темнее серого цвета**.</span><span class="sxs-lookup"><span data-stu-id="90d39-279">Text lines not in a container should be **Dark Gray**.</span></span>

### <a name="text"></a><span data-ttu-id="90d39-280">Текст</span><span class="sxs-lookup"><span data-stu-id="90d39-280">Text</span></span>

<span data-ttu-id="90d39-281">Избегайте использования текстовых символов в значках.</span><span class="sxs-lookup"><span data-stu-id="90d39-281">Avoid using text characters in icons.</span></span> <span data-ttu-id="90d39-282">Так как продукты Office используются по всему миру, мы хотим, чтобы значки были как можно более независящими от языка.</span><span class="sxs-lookup"><span data-stu-id="90d39-282">Since Office products are used around the world, we want to keep icons as language neutral as possible.</span></span>

## <a name="production"></a><span data-ttu-id="90d39-283">Производственная среда</span><span class="sxs-lookup"><span data-stu-id="90d39-283">Production</span></span>

### <a name="icon-file-format"></a><span data-ttu-id="90d39-284">Формат файлов значков</span><span class="sxs-lookup"><span data-stu-id="90d39-284">Icon file format</span></span>

<span data-ttu-id="90d39-285">Последние значки необходимо сохранить в виде PNG-файлов.</span><span class="sxs-lookup"><span data-stu-id="90d39-285">The final icons should be saved as .png image files.</span></span> <span data-ttu-id="90d39-286">Используйте формат PNG с прозрачным фоном и за32-разрядная глубина.</span><span class="sxs-lookup"><span data-stu-id="90d39-286">Use PNG format with a transparent background and have 32-bit depth.</span></span>
