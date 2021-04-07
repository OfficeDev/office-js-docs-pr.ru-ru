---
title: Руководство по значкам стилей Monoline для надстройок Office
description: Рекомендации по использованию значков стилей Monoline в надстройки Office.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: b74b89b2d622a6166fa111ef92bd8b2fffe79f8a
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604675"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a><span data-ttu-id="edf87-103">Руководство по значкам стилей Monoline для надстройок Office</span><span class="sxs-lookup"><span data-stu-id="edf87-103">Monoline style icon guidelines for Office Add-ins</span></span>

<span data-ttu-id="edf87-104">Иконография стилей Monoline используется в приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="edf87-104">Monoline style iconography are used in Office apps.</span></span> <span data-ttu-id="edf87-105">Если вы предпочитаете, чтобы значки совпадали со стилем Fresh для Office 2013+, см. рекомендации по значкам в стиле Fresh для [надстройок Office.](add-in-icons-fresh.md)</span><span class="sxs-lookup"><span data-stu-id="edf87-105">If you would prefer that your icons match the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

## <a name="office-monoline-visual-style"></a><span data-ttu-id="edf87-106">Визуальный стиль Office Monoline</span><span class="sxs-lookup"><span data-stu-id="edf87-106">Office Monoline visual style</span></span>

<span data-ttu-id="edf87-107">Цель стиля Monoline — иметь согласованную, четкую и доступную иконографию для связи действий и функций с помощью простых визуальных эффектов, обеспечения доступности значков для всех пользователей и стиля, соответствующего тем, которые используются в других частях Windows.</span><span class="sxs-lookup"><span data-stu-id="edf87-107">The goal of the Monoline style to have consistent, clear, and accessible iconography to communicate action and features with simple visuals, ensure the icons are accessible to all users, and have a style that is consistent with those used elsewhere in Windows.</span></span>

<span data-ttu-id="edf87-108">Ниже представлены рекомендации для сторонних разработчиков, которые хотят создать значки для функций, которые будут соответствовать значкам, уже представленным продуктам Office.</span><span class="sxs-lookup"><span data-stu-id="edf87-108">The following guidelines are for 3rd party developers who want to create icons for features that will be consistent with the icons already present Office products.</span></span>

### <a name="design-principles"></a><span data-ttu-id="edf87-109">Принципы разработки</span><span class="sxs-lookup"><span data-stu-id="edf87-109">Design principles</span></span>

- <span data-ttu-id="edf87-110">Простой, чистый, понятный.</span><span class="sxs-lookup"><span data-stu-id="edf87-110">Simple, clean, clear.</span></span>
- <span data-ttu-id="edf87-111">Содержит только необходимые элементы.</span><span class="sxs-lookup"><span data-stu-id="edf87-111">Contain only necessary elements.</span></span>
- <span data-ttu-id="edf87-112">Вдохновленный стилем значка Windows.</span><span class="sxs-lookup"><span data-stu-id="edf87-112">Inspired by Windows icon style.</span></span>
- <span data-ttu-id="edf87-113">Доступно для всех пользователей.</span><span class="sxs-lookup"><span data-stu-id="edf87-113">Accessible to all users.</span></span>

#### <a name="conveying-meaning"></a><span data-ttu-id="edf87-114">Передающее значение</span><span class="sxs-lookup"><span data-stu-id="edf87-114">Conveying meaning</span></span>

- <span data-ttu-id="edf87-115">Используйте описательные элементы, такие как страница, чтобы представлять документ или конверт для представления почты.</span><span class="sxs-lookup"><span data-stu-id="edf87-115">Use descriptive elements such as a page to represent a document or an envelope to represent mail.</span></span>
- <span data-ttu-id="edf87-116">Используйте один и тот же элемент для представления одной и той же концепции, то есть почта всегда представлена конвертом, а не штампом.</span><span class="sxs-lookup"><span data-stu-id="edf87-116">Use the same element to represent the same concept, i.e., mail is always represented by an envelope, not a stamp.</span></span>
- <span data-ttu-id="edf87-117">Используйте основную метафору во время разработки концепции.</span><span class="sxs-lookup"><span data-stu-id="edf87-117">Use a core metaphor during concept development.</span></span>

#### <a name="reduction-of-elements"></a><span data-ttu-id="edf87-118">Уменьшение элементов</span><span class="sxs-lookup"><span data-stu-id="edf87-118">Reduction of Elements</span></span>

- <span data-ttu-id="edf87-119">Уменьшите значок до основного значения, используя только элементы, необходимые для метафоры.</span><span class="sxs-lookup"><span data-stu-id="edf87-119">Reduce the icon to its core meaning, using only elements that are essential to the metaphor.</span></span>
- <span data-ttu-id="edf87-120">Ограничить число элементов в значке двумя, независимо от размера значка.</span><span class="sxs-lookup"><span data-stu-id="edf87-120">Limit the number of elements in an icon to two, regardless of icon size.</span></span>

#### <a name="consistency"></a><span data-ttu-id="edf87-121">Согласованность</span><span class="sxs-lookup"><span data-stu-id="edf87-121">Consistency</span></span>

<span data-ttu-id="edf87-122">Размеры, расположение и цвет значков должны быть последовательными.</span><span class="sxs-lookup"><span data-stu-id="edf87-122">Sizes, arrangement, and color of icons should be consistent.</span></span>

#### <a name="styling"></a><span data-ttu-id="edf87-123">Стиль</span><span class="sxs-lookup"><span data-stu-id="edf87-123">Styling</span></span>

##### <a name="perspective"></a><span data-ttu-id="edf87-124">Perspective</span><span class="sxs-lookup"><span data-stu-id="edf87-124">Perspective</span></span>

<span data-ttu-id="edf87-125">Значки Monoline по умолчанию имеют передовую линию.</span><span class="sxs-lookup"><span data-stu-id="edf87-125">Monoline icons are forward-facing by default.</span></span> <span data-ttu-id="edf87-126">Допускаются некоторые элементы, которые требуют перспективы и/или вращения, например куба, но исключения следует совмещение.</span><span class="sxs-lookup"><span data-stu-id="edf87-126">Certain elements that require perspective and/or rotation, such as a cube, are allowed, but exceptions should be kept to a minimum.</span></span>

##### <a name="embellishment"></a><span data-ttu-id="edf87-127">Приукрашивание</span><span class="sxs-lookup"><span data-stu-id="edf87-127">Embellishment</span></span>

<span data-ttu-id="edf87-128">Monoline — это чистый минимальный стиль.</span><span class="sxs-lookup"><span data-stu-id="edf87-128">Monoline is a clean minimal style.</span></span> <span data-ttu-id="edf87-129">Все использует плоский цвет, что означает отсутствие градиентов, текстур или источников света.</span><span class="sxs-lookup"><span data-stu-id="edf87-129">Everything uses flat color, which means there are no gradients, textures, or light sources.</span></span>

## <a name="designing"></a><span data-ttu-id="edf87-130">Проектирование</span><span class="sxs-lookup"><span data-stu-id="edf87-130">Designing</span></span>

### <a name="sizes"></a><span data-ttu-id="edf87-131">Размеры</span><span class="sxs-lookup"><span data-stu-id="edf87-131">Sizes</span></span>

<span data-ttu-id="edf87-132">Мы рекомендуем создавать каждый значок всех этих размеров для поддержки устройств с высоким уровнем DPI.</span><span class="sxs-lookup"><span data-stu-id="edf87-132">We recommend that you produce each icon in all these sizes to support high DPI devices.</span></span> <span data-ttu-id="edf87-133">Абсолютно *необходимые* размеры 16 px, 20 px и 32 px, так как это 100% размеров.</span><span class="sxs-lookup"><span data-stu-id="edf87-133">The absolutely *required* sizes are 16 px, 20 px, and 32 px, as those are the 100% sizes.</span></span>

<span data-ttu-id="edf87-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span><span class="sxs-lookup"><span data-stu-id="edf87-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span></span>

> [!IMPORTANT]
> <span data-ttu-id="edf87-135">Для изображения, которое является символом представительства надстройки, см. в статью Создание эффективных списков в [AppSource](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) и Office для размера и других требований.</span><span class="sxs-lookup"><span data-stu-id="edf87-135">For an image that is your add-in's representative icon, see [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in) for size and other requirements.</span></span>

### <a name="layout"></a><span data-ttu-id="edf87-136">Макет</span><span class="sxs-lookup"><span data-stu-id="edf87-136">Layout</span></span>

<span data-ttu-id="edf87-137">Ниже приводится пример макета значков с модификатором.</span><span class="sxs-lookup"><span data-stu-id="edf87-137">The following is an example of icon layout with a modifier.</span></span>

![Схема значка с модификатором в правом нижнем ряду](../images/monolineicon1.png)  ![Схема одной и той же значки с добавленным фоном сетки и вызовами для базы, модификатора, обивки и выреза](../images/monolineicon2.png)

#### <a name="elements"></a><span data-ttu-id="edf87-140">Элементы</span><span class="sxs-lookup"><span data-stu-id="edf87-140">Elements</span></span>

- <span data-ttu-id="edf87-141">**База.** Основная концепция, которую представляет значок.</span><span class="sxs-lookup"><span data-stu-id="edf87-141">**Base**: The main concept that the icon represents.</span></span> <span data-ttu-id="edf87-142">Обычно это единственный визуальный элемент, необходимый для значка, но иногда основная концепция может быть улучшена с помощью дополнительного элемента, модификатора.</span><span class="sxs-lookup"><span data-stu-id="edf87-142">This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.</span></span>

- <span data-ttu-id="edf87-143">**Модификатор** Любой элемент, наложение на базу; то есть модификатор, который обычно представляет действие или состояние.</span><span class="sxs-lookup"><span data-stu-id="edf87-143">**Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status.</span></span> <span data-ttu-id="edf87-144">Он изменяет базовый элемент, выступая в качестве добавления, изменения или дескриптора.</span><span class="sxs-lookup"><span data-stu-id="edf87-144">It modifies the base element by acting as an addition, alteration, or a descriptor.</span></span>

![Схема сетки с областями базовых и модификаторов](../images/monolineicon3.png)

### <a name="construction"></a><span data-ttu-id="edf87-146">Строительство</span><span class="sxs-lookup"><span data-stu-id="edf87-146">Construction</span></span>

#### <a name="element-placement"></a><span data-ttu-id="edf87-147">Размещение элементов</span><span class="sxs-lookup"><span data-stu-id="edf87-147">Element placement</span></span>

<span data-ttu-id="edf87-148">Базовые элементы размещаются в центре значка в обивке.</span><span class="sxs-lookup"><span data-stu-id="edf87-148">Base elements are placed in the center of the icon within the padding.</span></span> <span data-ttu-id="edf87-149">Если он не может быть размещен идеально центр, то база должна err в правом верхнем.</span><span class="sxs-lookup"><span data-stu-id="edf87-149">If it can't be placed perfectly centered, then the base should err to the top right.</span></span> <span data-ttu-id="edf87-150">В следующем примере иконка идеально центризолась.</span><span class="sxs-lookup"><span data-stu-id="edf87-150">In the following example, the icon is perfectly centered.</span></span>

![Диаграмма, показывающая идеально центру значок](../images/monolineicon4.png)

<span data-ttu-id="edf87-152">В следующем примере значок забвещается влево.</span><span class="sxs-lookup"><span data-stu-id="edf87-152">In the following example, the icon is erring to the left.</span></span>

![Диаграмма, показывающая значок, который перебегает слева на 1 px](../images/monolineicon5.png)

<span data-ttu-id="edf87-154">Модификаторы почти всегда размещаются в правом нижнем углу холста значков.</span><span class="sxs-lookup"><span data-stu-id="edf87-154">Modifiers are almost always placed in the bottom right corner of the icon canvas.</span></span> <span data-ttu-id="edf87-155">В некоторых редких случаях модификаторы помещаются в другой угол.</span><span class="sxs-lookup"><span data-stu-id="edf87-155">In some rare cases, modifiers are placed in a different corner.</span></span> <span data-ttu-id="edf87-156">Например, если базовый элемент будет неузнаваем с модификатором в правом нижнем углу, то рассмотрите возможность его размещения в верхнем левом углу.</span><span class="sxs-lookup"><span data-stu-id="edf87-156">For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.</span></span>

![Схема, показывающая четыре значка с модификатором в правом нижнем ряду и один значок с модификатором в верхнем левом.](../images/monolineicon6.png)

#### <a name="padding"></a><span data-ttu-id="edf87-158">Обивка</span><span class="sxs-lookup"><span data-stu-id="edf87-158">Padding</span></span>

<span data-ttu-id="edf87-159">Каждый значок размера имеет определенное количество обивки вокруг значка.</span><span class="sxs-lookup"><span data-stu-id="edf87-159">Each size icon has a specified amount of padding around the icon.</span></span> <span data-ttu-id="edf87-160">Базовый элемент остается в области обивки, но модификатор должен прикладом до края холста, простираясь за пределы обивки до края границы значка.</span><span class="sxs-lookup"><span data-stu-id="edf87-160">The base element stays within the padding, but the modifier should butt up to the edge of the canvas, extending outside of the padding to the edge of the icon border.</span></span> <span data-ttu-id="edf87-161">На следующих изображениях покажите рекомендуемую обивку для каждого из размеров значка.</span><span class="sxs-lookup"><span data-stu-id="edf87-161">The following images show the recommended padding to use for each of the icon sizes.</span></span>

|<span data-ttu-id="edf87-162">**16 пк**</span><span class="sxs-lookup"><span data-stu-id="edf87-162">**16px**</span></span>|<span data-ttu-id="edf87-163">**20 пк**</span><span class="sxs-lookup"><span data-stu-id="edf87-163">**20px**</span></span>|<span data-ttu-id="edf87-164">**24 пк**</span><span class="sxs-lookup"><span data-stu-id="edf87-164">**24px**</span></span>|<span data-ttu-id="edf87-165">**32 пк**</span><span class="sxs-lookup"><span data-stu-id="edf87-165">**32px**</span></span>|<span data-ttu-id="edf87-166">**40 пк**</span><span class="sxs-lookup"><span data-stu-id="edf87-166">**40px**</span></span>|<span data-ttu-id="edf87-167">**48 пк**</span><span class="sxs-lookup"><span data-stu-id="edf87-167">**48px**</span></span>|<span data-ttu-id="edf87-168">**64 пк**</span><span class="sxs-lookup"><span data-stu-id="edf87-168">**64px**</span></span>|<span data-ttu-id="edf87-169">**80 пк**</span><span class="sxs-lookup"><span data-stu-id="edf87-169">**80px**</span></span>|<span data-ttu-id="edf87-170">**96px**</span><span class="sxs-lookup"><span data-stu-id="edf87-170">**96px**</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![Значок 16 px с обивкой 0px](../images/monolineicon7.png)|![Значок 20 px с обивкой 1px](../images/monolineicon8.png)|![Значок 24 px с обивкой 1px](../images/monolineicon9.png)|![Значок 32 px с обивкой 2px](../images/monolineicon10.png)|![Значок 40 px с обивкой 2px](../images/monolineicon11.png)|![Значок 48 px с обивкой 3px](../images/monolineicon12.png)|![Значок 64 px с обивкой 4px](../images/monolineicon13.png)|![Значок 80 px с обивкой 5px](../images/monolineicon14.png)|![Значок 96 px с обивкой 6px](../images/monolineicon15.png)|

#### <a name="line-weights"></a><span data-ttu-id="edf87-180">Весы строки</span><span class="sxs-lookup"><span data-stu-id="edf87-180">Line weights</span></span>

<span data-ttu-id="edf87-181">Monoline — это стиль, в котором преобладают линии и контурные фигуры.</span><span class="sxs-lookup"><span data-stu-id="edf87-181">Monoline is a style dominated by line and outlined shapes.</span></span> <span data-ttu-id="edf87-182">В зависимости от размера, который вы производите, значок должен использовать следующие весы строки.</span><span class="sxs-lookup"><span data-stu-id="edf87-182">Depending on what size you are producing the icon should use the following line weights.</span></span>

|<span data-ttu-id="edf87-183">Размер значка:</span><span class="sxs-lookup"><span data-stu-id="edf87-183">Icon Size:</span></span>|<span data-ttu-id="edf87-184">16 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-184">16px</span></span>|<span data-ttu-id="edf87-185">20 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-185">20px</span></span>|<span data-ttu-id="edf87-186">24 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-186">24px</span></span>|<span data-ttu-id="edf87-187">32 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-187">32px</span></span>|<span data-ttu-id="edf87-188">40 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-188">40px</span></span>|<span data-ttu-id="edf87-189">48 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-189">48px</span></span>|<span data-ttu-id="edf87-190">64 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-190">64px</span></span>|<span data-ttu-id="edf87-191">80 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-191">80px</span></span>|<span data-ttu-id="edf87-192">96px</span><span class="sxs-lookup"><span data-stu-id="edf87-192">96px</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|<span data-ttu-id="edf87-193">**Вес строки:**</span><span class="sxs-lookup"><span data-stu-id="edf87-193">**Line Weight:**</span></span>|<span data-ttu-id="edf87-194">1 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-194">1px</span></span>|<span data-ttu-id="edf87-195">1 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-195">1px</span></span>|<span data-ttu-id="edf87-196">1 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-196">1px</span></span>|<span data-ttu-id="edf87-197">1 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-197">1px</span></span>|<span data-ttu-id="edf87-198">2 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-198">2px</span></span>|<span data-ttu-id="edf87-199">2 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-199">2px</span></span>|<span data-ttu-id="edf87-200">2 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-200">2px</span></span>|<span data-ttu-id="edf87-201">2 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-201">2px</span></span>|<span data-ttu-id="edf87-202">3 пк</span><span class="sxs-lookup"><span data-stu-id="edf87-202">3px</span></span>|
|<span data-ttu-id="edf87-203">**Значок примера:**</span><span class="sxs-lookup"><span data-stu-id="edf87-203">**Example icon:**</span></span>|![Значок 16 px](../images/monolineicon16.png)|![Значок 20 px](../images/monolineicon17.png)|![Значок 24 px](../images/monolineicon18.png)|![Значок 32 px](../images/monolineicon19.png)|![Значок 40 px](../images/monolineicon20.png)|![Значок 48 px](../images/monolineicon21.png)|![Значок 64 px](../images/monolineicon22.png)|![Значок 80 px](../images/monolineicon23.png)|![Значок 96 px](../images/monolineicon24.png)|

#### <a name="cutouts"></a><span data-ttu-id="edf87-213">Вырезы</span><span class="sxs-lookup"><span data-stu-id="edf87-213">Cutouts</span></span>

<span data-ttu-id="edf87-214">Когда элемент значка помещается поверх другого элемента, вырез (нижнего элемента) используется для предоставления пространства между двумя элементами, главным образом для целей читаемости.</span><span class="sxs-lookup"><span data-stu-id="edf87-214">When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes.</span></span> <span data-ttu-id="edf87-215">Это обычно происходит, когда модификатор помещается поверх базового элемента, но есть также случаи, когда ни один из элементов не является модификатором.</span><span class="sxs-lookup"><span data-stu-id="edf87-215">This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier.</span></span> <span data-ttu-id="edf87-216">Эти вырезы между двумя элементами иногда называются "разрывом".</span><span class="sxs-lookup"><span data-stu-id="edf87-216">These cutouts between the two elements is sometimes referred to as a "gap".</span></span>

<span data-ttu-id="edf87-217">Размер зазора должен быть такой же шириной, как и вес строки, используемый для этого размера.</span><span class="sxs-lookup"><span data-stu-id="edf87-217">The size of the gap should be the same width as the line weight used on that size.</span></span> <span data-ttu-id="edf87-218">Если сделать значок 16 px, ширина пробела будет 1px, а если это значок 48 px, то разрыв должен быть 2px.</span><span class="sxs-lookup"><span data-stu-id="edf87-218">If making a 16 px icon, the gap width would be 1px and if it is a 48 px icon then the gap should be 2px.</span></span> <span data-ttu-id="edf87-219">В следующем примере показан значок 32 px с разрывом в 1px между модификатором и базовой базой.</span><span class="sxs-lookup"><span data-stu-id="edf87-219">The following example shows a 32 px icon with a gap of 1px between the modifier and the underlying base.</span></span>

![Значок 32 px с разрывом 1px между модификатором и базовой базой](../images/monolineicon25.png)

<span data-ttu-id="edf87-221">В некоторых случаях разрыв может быть увеличен на 1/2 px, если модификатор имеет диагональный или изогнутый край, а стандартный разрыв не обеспечивает достаточного разделения.</span><span class="sxs-lookup"><span data-stu-id="edf87-221">In some cases, the gap can be increase by a 1/2 px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation.</span></span> <span data-ttu-id="edf87-222">Это, скорее всего, повлияет только на значки с весом строки 1px: 16 px, 20 px, 24 px и 32 px.</span><span class="sxs-lookup"><span data-stu-id="edf87-222">This will likely only affect the icons with 1px line weight: 16 px, 20 px, 24 px, and 32 px.</span></span>

#### <a name="background-fills"></a><span data-ttu-id="edf87-223">Заполнение фона</span><span class="sxs-lookup"><span data-stu-id="edf87-223">Background fills</span></span>

<span data-ttu-id="edf87-224">Большинство значков в наборе значков Monoline требуют заполнения фона.</span><span class="sxs-lookup"><span data-stu-id="edf87-224">Most icons in the Monoline icon set require background fills.</span></span> <span data-ttu-id="edf87-225">Однако существуют случаи, когда объект не имеет естественного заполнения, поэтому не следует применять заливки.</span><span class="sxs-lookup"><span data-stu-id="edf87-225">However, there are cases where the object would not naturally have a fill, so no fill should be applied.</span></span> <span data-ttu-id="edf87-226">Следующие значки имеют белый заливки.</span><span class="sxs-lookup"><span data-stu-id="edf87-226">The following icons have a white fill.</span></span>

![Компиляция пяти значков с белым заливом](../images/monolineicon26.png)

<span data-ttu-id="edf87-228">Следующие значки не заполняются.</span><span class="sxs-lookup"><span data-stu-id="edf87-228">The following icons have no fill.</span></span> <span data-ttu-id="edf87-229">(Значок передач включен, чтобы показать, что центральное отверстие не заполнено.)</span><span class="sxs-lookup"><span data-stu-id="edf87-229">(The gear icon is included to show that the center hole is not filled.)</span></span>

![Компиляция пяти значков без заполнения](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a><span data-ttu-id="edf87-231">Лучшие практики для заполнения</span><span class="sxs-lookup"><span data-stu-id="edf87-231">Best practices for fills</span></span>

###### <a name="dos"></a><span data-ttu-id="edf87-232">Dos:</span><span class="sxs-lookup"><span data-stu-id="edf87-232">Dos:</span></span>

- <span data-ttu-id="edf87-233">Заполните любой элемент с определенной границей и естественным образом заполните его.</span><span class="sxs-lookup"><span data-stu-id="edf87-233">Fill any element that has a defined boundary, and would naturally have a fill.</span></span>
- <span data-ttu-id="edf87-234">Для создания фонового заполнения используйте отдельную форму.</span><span class="sxs-lookup"><span data-stu-id="edf87-234">Use a separate shape to create the background fill.</span></span>
- <span data-ttu-id="edf87-235">Используйте **фоновое заполнение** [из цветовой палитры](#color).</span><span class="sxs-lookup"><span data-stu-id="edf87-235">Use **Background Fill** from the [color palette](#color).</span></span>
- <span data-ttu-id="edf87-236">Сохранение разделения пикселей между перекрывающимися элементами.</span><span class="sxs-lookup"><span data-stu-id="edf87-236">Maintain the pixel separation between overlapping elements.</span></span>
- <span data-ttu-id="edf87-237">Заполните между несколькими объектами.</span><span class="sxs-lookup"><span data-stu-id="edf87-237">Fill between multiple objects.</span></span>

###### <a name="donts"></a><span data-ttu-id="edf87-238">Не нужно:</span><span class="sxs-lookup"><span data-stu-id="edf87-238">Don'ts:</span></span>

- <span data-ttu-id="edf87-239">Не заполняйте объекты, которые естественно не заполняются; например, сальто.</span><span class="sxs-lookup"><span data-stu-id="edf87-239">Don't fill objects that would not naturally be filled; for example, a paperclip.</span></span>
- <span data-ttu-id="edf87-240">Не заполняйте скобки.</span><span class="sxs-lookup"><span data-stu-id="edf87-240">Don't fill brackets.</span></span>
- <span data-ttu-id="edf87-241">Не заполняйте номера или альфа-символы.</span><span class="sxs-lookup"><span data-stu-id="edf87-241">Don't fill behind numbers or alpha characters.</span></span>

### <a name="color"></a><span data-ttu-id="edf87-242">Цвет</span><span class="sxs-lookup"><span data-stu-id="edf87-242">Color</span></span>

<span data-ttu-id="edf87-243">Цветовая палитра разработана для простоты и доступности.</span><span class="sxs-lookup"><span data-stu-id="edf87-243">The color palette has been designed for simplicity and accessibility.</span></span> <span data-ttu-id="edf87-244">Он содержит 4 нейтральных цвета и два варианта для синего, зеленого, желтого, красного и фиолетового цветов.</span><span class="sxs-lookup"><span data-stu-id="edf87-244">It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple.</span></span> <span data-ttu-id="edf87-245">Оранжевый цвет намеренно не входит в цветовую палитру значков Monoline.</span><span class="sxs-lookup"><span data-stu-id="edf87-245">Orange is intentionally not included in the Monoline icon color palette.</span></span> <span data-ttu-id="edf87-246">Каждый цвет предназначен для использования определенными способами, как описано в этом разделе.</span><span class="sxs-lookup"><span data-stu-id="edf87-246">Each color is intended to be used in specific ways as outlined in this section.</span></span>

#### <a name="palette"></a><span data-ttu-id="edf87-247">Палитра</span><span class="sxs-lookup"><span data-stu-id="edf87-247">Palette</span></span>

![Четыре оттенка серого в монолайне: темно-серый для автономных или контурных, средний серый для контура или контента, очень светло-серый для заполнения фона и светло-серый для заполнения](../images/monoline-grayshades.png)

![Цветовая палитра в монолине включает в себя оттенок синего, зеленого, желтого, красного и фиолетового для автономных, контурных и заполняемых](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a><span data-ttu-id="edf87-250">Использование цвета</span><span class="sxs-lookup"><span data-stu-id="edf87-250">How to use color</span></span>

<span data-ttu-id="edf87-251">В цветовой палитре Monoline все цвета имеют автономные, контурные и заполняемые варианты.</span><span class="sxs-lookup"><span data-stu-id="edf87-251">In the Monoline color palette, all colors have Standalone, Outline, and Fill variations.</span></span> <span data-ttu-id="edf87-252">Как правило, элементы построены с заполнием и границей.</span><span class="sxs-lookup"><span data-stu-id="edf87-252">Generally, elements are constructed with a fill and a border.</span></span> <span data-ttu-id="edf87-253">Цвета применяются в одном из следующих шаблонов:</span><span class="sxs-lookup"><span data-stu-id="edf87-253">The colors are applied in one of the following patterns:</span></span>

- <span data-ttu-id="edf87-254">Автономный цвет для объектов без заполнения.</span><span class="sxs-lookup"><span data-stu-id="edf87-254">The Standalone color alone for objects that have no fill.</span></span>
- <span data-ttu-id="edf87-255">На границе используется цвет Outline, а для заполнения используется цвет Fill.</span><span class="sxs-lookup"><span data-stu-id="edf87-255">The border uses the Outline color and the fill uses the Fill color.</span></span>
- <span data-ttu-id="edf87-256">На границе используется автономный цвет, а для заполнения используется цвет Фоновое заполнение.</span><span class="sxs-lookup"><span data-stu-id="edf87-256">The border uses the Standalone color and the fill uses the Background Fill color.</span></span>

<span data-ttu-id="edf87-257">Ниже приводится пример использования цвета.</span><span class="sxs-lookup"><span data-stu-id="edf87-257">The following are examples of using color.</span></span>

![Компиляция трех значков с цветом на границе или заполнения или обоих](../images/monolineicon28.png)

<span data-ttu-id="edf87-259">Наиболее распространенной ситуацией будет использование элемента Темно-серый автономный с фоновой заливки.</span><span class="sxs-lookup"><span data-stu-id="edf87-259">The most common situation will be to have an element use Dark Gray Standalone with Background Fill.</span></span>

<span data-ttu-id="edf87-260">При использовании цветной заливки он всегда должен быть с соответствующим цветом Outline.</span><span class="sxs-lookup"><span data-stu-id="edf87-260">When using a colored Fill, it should always be with its corresponding Outline color.</span></span> <span data-ttu-id="edf87-261">Например, Blue Fill следует использовать только с помощью Blue Outline.</span><span class="sxs-lookup"><span data-stu-id="edf87-261">For example, Blue Fill should only be used with Blue Outline.</span></span> <span data-ttu-id="edf87-262">Но есть два исключения из этого общего правила:</span><span class="sxs-lookup"><span data-stu-id="edf87-262">But there are two exceptions to this general rule:</span></span>

- <span data-ttu-id="edf87-263">Фоновое заполнение можно использовать с любым автономным цветом.</span><span class="sxs-lookup"><span data-stu-id="edf87-263">Background Fill can be used with any color Standalone.</span></span>
- <span data-ttu-id="edf87-264">Светло-серый заливка может использоваться с двумя разными цветами Outline: темно-серый или средний серый.</span><span class="sxs-lookup"><span data-stu-id="edf87-264">Light Gray Fill can be used with two different Outline colors: Dark Gray or Medium Gray.</span></span>

#### <a name="when-to-use-color"></a><span data-ttu-id="edf87-265">Когда использовать цвет</span><span class="sxs-lookup"><span data-stu-id="edf87-265">When to use color</span></span>

<span data-ttu-id="edf87-266">Цвет должен использоваться для передачи значения значка, а не для украшения.</span><span class="sxs-lookup"><span data-stu-id="edf87-266">Color should be used to convey the meaning of the icon rather than for embellishment.</span></span> <span data-ttu-id="edf87-267">Он должен **выделить действие** пользователю.</span><span class="sxs-lookup"><span data-stu-id="edf87-267">It should **highlight the action** to the user.</span></span> <span data-ttu-id="edf87-268">При добавлении модификатора в базовый элемент с цветом базовый элемент обычно превращается в темно-серый и фоновый, так что модификатор может быть элементом цвета, например в случае ниже с модификатором "X", который добавляется в базу изображений в левом значке следующего набора.</span><span class="sxs-lookup"><span data-stu-id="edf87-268">When a modifier is added to a base element that has color, the base element is typically turned into Dark Gray and Background Fill so that the modifier can be the element of color, such as the case below with the "X" modifier being added to the picture base in the leftmost icon of the following set.</span></span>

![Компиляция пяти значков с использованием цвета](../images/monolineicon29.png)

<span data-ttu-id="edf87-270">Вы должны ограничить значки **одним дополнительным** цветом, кроме описанных выше набросков и заливки.</span><span class="sxs-lookup"><span data-stu-id="edf87-270">You should limit your icons to **one** additional color, other than the Outline and Fill mentioned above.</span></span> <span data-ttu-id="edf87-271">Тем не менее, больше цветов можно использовать, если это жизненно важно для его метафоры, с ограничением двух дополнительных цветов, кроме серого.</span><span class="sxs-lookup"><span data-stu-id="edf87-271">However, more colors can be used if it is vital for its metaphor, with a limit of two additional colors other than gray.</span></span> <span data-ttu-id="edf87-272">В редких случаях существуют исключения, когда требуется больше цветов.</span><span class="sxs-lookup"><span data-stu-id="edf87-272">In rare cases, there are exceptions when more colors are needed.</span></span> <span data-ttu-id="edf87-273">Ниже приводится хороший пример значков, которые используют только один цвет.</span><span class="sxs-lookup"><span data-stu-id="edf87-273">The following are good examples of icons that use just one color.</span></span>

  ![Компиляция пяти значков, каждый из которых использует один цвет](../images/monolineicon30.png)

<span data-ttu-id="edf87-275">Но в следующих значках используется слишком много цветов.</span><span class="sxs-lookup"><span data-stu-id="edf87-275">But the following icons use too many colors.</span></span>

  ![Компиляция пяти значков, каждый из которых использует несколько цветов](../images/monolineicon31.png)

<span data-ttu-id="edf87-277">Используйте **medium Gray** для внутреннего "контента", например линий сетки в значке таблицы.</span><span class="sxs-lookup"><span data-stu-id="edf87-277">Use **Medium Gray** for interior "content", such as grid lines in an icon of a spreadsheet.</span></span> <span data-ttu-id="edf87-278">Дополнительные цвета интерьера используются, когда содержимое должно показывать поведение управления.</span><span class="sxs-lookup"><span data-stu-id="edf87-278">Additional interior colors are used when the content needs to show the behavior of the control.</span></span>

![Компиляция пяти значков со средними серыми элементами интерьера](../images/monolineicon32.png)

#### <a name="text-lines"></a><span data-ttu-id="edf87-280">Текстовые строки</span><span class="sxs-lookup"><span data-stu-id="edf87-280">Text lines</span></span>

<span data-ttu-id="edf87-281">Если текстовые строки находятся в "контейнере" (например, текст на документе), используйте средне-серый цвет.</span><span class="sxs-lookup"><span data-stu-id="edf87-281">When text lines are in a "container" (for example, text on a document), use medium gray.</span></span> <span data-ttu-id="edf87-282">Текстовые строки, не в контейнере, должны быть **темно-серыми.**</span><span class="sxs-lookup"><span data-stu-id="edf87-282">Text lines not in a container should be **Dark Gray**.</span></span>

### <a name="text"></a><span data-ttu-id="edf87-283">Текст</span><span class="sxs-lookup"><span data-stu-id="edf87-283">Text</span></span>

<span data-ttu-id="edf87-284">Избегайте использования текстовых символов в значках.</span><span class="sxs-lookup"><span data-stu-id="edf87-284">Avoid using text characters in icons.</span></span> <span data-ttu-id="edf87-285">Так как продукты Office используются по всему миру, мы хотим сохранить значки как можно более нейтральными на языке.</span><span class="sxs-lookup"><span data-stu-id="edf87-285">Since Office products are used around the world, we want to keep icons as language neutral as possible.</span></span>

## <a name="production"></a><span data-ttu-id="edf87-286">Производство</span><span class="sxs-lookup"><span data-stu-id="edf87-286">Production</span></span>

### <a name="icon-file-format"></a><span data-ttu-id="edf87-287">Формат файла icon</span><span class="sxs-lookup"><span data-stu-id="edf87-287">Icon file format</span></span>

<span data-ttu-id="edf87-288">Конечные значки должны быть сохранены в качестве файлов изображений png.</span><span class="sxs-lookup"><span data-stu-id="edf87-288">The final icons should be saved as .png image files.</span></span> <span data-ttu-id="edf87-289">Используйте формат PNG с прозрачным фоном и 32-битной глубиной.</span><span class="sxs-lookup"><span data-stu-id="edf87-289">Use PNG format with a transparent background and have 32-bit depth.</span></span>

## <a name="see-also"></a><span data-ttu-id="edf87-290">См. также</span><span class="sxs-lookup"><span data-stu-id="edf87-290">See also</span></span>

- [<span data-ttu-id="edf87-291">Элемент манифеста Icon</span><span class="sxs-lookup"><span data-stu-id="edf87-291">Icon manifest element</span></span>](../reference/manifest/icon.md)
- [<span data-ttu-id="edf87-292">Элемент манифеста IconUrl</span><span class="sxs-lookup"><span data-stu-id="edf87-292">IconUrl manifest element</span></span>](../reference/manifest/iconurl.md)
- [<span data-ttu-id="edf87-293">Элемент манифеста HighResolutionIconUrl</span><span class="sxs-lookup"><span data-stu-id="edf87-293">HighResolutionIconUrl manifest element</span></span>](../reference/manifest/highresolutioniconurl.md)
- [<span data-ttu-id="edf87-294">Создание значка для надстройки</span><span class="sxs-lookup"><span data-stu-id="edf87-294">Create an icon for your add-in</span></span>](/office/dev/store/create-effective-office-store-listings#create-an-icon-for-your-add-in)
