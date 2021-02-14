---
title: Руководство по значкам моностройного стиля для надстройки Office
description: Получите рекомендации по использованию значков стилей Monoline в надстройки Office.
ms.date: 2/09/2021
localization_priority: Normal
ms.openlocfilehash: 262cde129c7f7d3dd3f32b32e0a8e750cf016ef8
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237954"
---
# <a name="monoline-style-icon-guidelines-for-office-add-ins"></a><span data-ttu-id="9d977-103">Руководство по значкам моностройного стиля для надстройки Office</span><span class="sxs-lookup"><span data-stu-id="9d977-103">Monoline style icon guidelines for Office Add-ins</span></span>

<span data-ttu-id="9d977-104">В приложениях Office используются монолинейные значки стилей.</span><span class="sxs-lookup"><span data-stu-id="9d977-104">Monoline style iconography are used in Office apps.</span></span> <span data-ttu-id="9d977-105">Если вы предпочитаете, чтобы ваши значки совпадали со стилем "Новый" office 2013 и office 2013 и более новыми, см. рекомендации по значкам "Новый стиль" для надстройки [Office.](add-in-icons-fresh.md)</span><span class="sxs-lookup"><span data-stu-id="9d977-105">If you would prefer that your icons match the Fresh style of non-subscription Office 2013+, see [Fresh style icon guidelines for Office Add-ins](add-in-icons-fresh.md).</span></span>

## <a name="office-monoline-visual-style"></a><span data-ttu-id="9d977-106">Визуальный стиль монолайна Office</span><span class="sxs-lookup"><span data-stu-id="9d977-106">Office Monoline visual style</span></span>

<span data-ttu-id="9d977-107">Цель стиля Monoline — обеспечить единообразие, четкое и доступное оформление для взаимодействия действий и функций с помощью простых визуальных объектов, обеспечить доступность значков для всех пользователей и иметь стиль, соответствующий стилю, который используется в других частях Windows.</span><span class="sxs-lookup"><span data-stu-id="9d977-107">The goal of the Monoline style to have consistent, clear, and accessible iconography to communicate action and features with simple visuals, ensure the icons are accessible to all users, and have a style that is consistent with those used elsewhere in Windows.</span></span>

<span data-ttu-id="9d977-108">Ниже представлены рекомендации для сторонних разработчиков, которые хотят создавать значки для функций, которые будут соответствовать значкам, уже представленным в продуктах Office.</span><span class="sxs-lookup"><span data-stu-id="9d977-108">The following guidelines are for 3rd party developers who want to create icons for features that will be consistent with the icons already present Office products.</span></span>

### <a name="design-principles"></a><span data-ttu-id="9d977-109">Принципы разработки</span><span class="sxs-lookup"><span data-stu-id="9d977-109">Design principles</span></span>

- <span data-ttu-id="9d977-110">Простое, чистое, понятное.</span><span class="sxs-lookup"><span data-stu-id="9d977-110">Simple, clean, clear.</span></span>
- <span data-ttu-id="9d977-111">Содержит только необходимые элементы.</span><span class="sxs-lookup"><span data-stu-id="9d977-111">Contain only necessary elements.</span></span>
- <span data-ttu-id="9d977-112">Навеяно стилем значка Windows.</span><span class="sxs-lookup"><span data-stu-id="9d977-112">Inspired by Windows icon style.</span></span>
- <span data-ttu-id="9d977-113">Доступно для всех пользователей.</span><span class="sxs-lookup"><span data-stu-id="9d977-113">Accessible to all users.</span></span>

#### <a name="conveying-meaning"></a><span data-ttu-id="9d977-114">Передав значение</span><span class="sxs-lookup"><span data-stu-id="9d977-114">Conveying meaning</span></span>

- <span data-ttu-id="9d977-115">Используйте описательные элементы, например страницу, для представления документа или конверта для представления почты.</span><span class="sxs-lookup"><span data-stu-id="9d977-115">Use descriptive elements such as a page to represent a document or an envelope to represent mail.</span></span>
- <span data-ttu-id="9d977-116">Используйте один и тот же элемент для представления одной и той же концепции, то есть почта всегда представлена конвертом, а не отметкой.</span><span class="sxs-lookup"><span data-stu-id="9d977-116">Use the same element to represent the same concept, i.e., mail is always represented by an envelope, not a stamp.</span></span>
- <span data-ttu-id="9d977-117">Используйте основную метафору во время разработки концепции.</span><span class="sxs-lookup"><span data-stu-id="9d977-117">Use a core metaphor during concept development.</span></span>

#### <a name="reduction-of-elements"></a><span data-ttu-id="9d977-118">Сокращение элементов</span><span class="sxs-lookup"><span data-stu-id="9d977-118">Reduction of Elements</span></span>

- <span data-ttu-id="9d977-119">Уменьшите значение значка до его основного значения, используя только элементы, необходимые для метафоры.</span><span class="sxs-lookup"><span data-stu-id="9d977-119">Reduce the icon to its core meaning, using only elements that are essential to the metaphor.</span></span>
- <span data-ttu-id="9d977-120">Ограничив количество элементов значка двумя, независимо от размера значка.</span><span class="sxs-lookup"><span data-stu-id="9d977-120">Limit the number of elements in an icon to two, regardless of icon size.</span></span>

#### <a name="consistency"></a><span data-ttu-id="9d977-121">Согласованность</span><span class="sxs-lookup"><span data-stu-id="9d977-121">Consistency</span></span>

<span data-ttu-id="9d977-122">Размеры, расположение и цвет значков должны быть согласованы.</span><span class="sxs-lookup"><span data-stu-id="9d977-122">Sizes, arrangement, and color of icons should be consistent.</span></span>

#### <a name="styling"></a><span data-ttu-id="9d977-123">Стиль</span><span class="sxs-lookup"><span data-stu-id="9d977-123">Styling</span></span>

##### <a name="perspective"></a><span data-ttu-id="9d977-124">Perspective</span><span class="sxs-lookup"><span data-stu-id="9d977-124">Perspective</span></span>

<span data-ttu-id="9d977-125">По умолчанию монолинейные значки перенадвигаются вперед.</span><span class="sxs-lookup"><span data-stu-id="9d977-125">Monoline icons are forward-facing by default.</span></span> <span data-ttu-id="9d977-126">Некоторые элементы, которые требуют перспективы и/или поворота, например куб, разрешены, но исключения должны быть с минимальными.</span><span class="sxs-lookup"><span data-stu-id="9d977-126">Certain elements that require perspective and/or rotation, such as a cube, are allowed, but exceptions should be kept to a minimum.</span></span>

##### <a name="embellishment"></a><span data-ttu-id="9d977-127">Безвластие</span><span class="sxs-lookup"><span data-stu-id="9d977-127">Embellishment</span></span>

<span data-ttu-id="9d977-128">Монолайн — это чистый минимальный стиль.</span><span class="sxs-lookup"><span data-stu-id="9d977-128">Monoline is a clean minimal style.</span></span> <span data-ttu-id="9d977-129">Все использует плоский цвет, то есть градиенты, текстуры и источники света не существуют.</span><span class="sxs-lookup"><span data-stu-id="9d977-129">Everything uses flat color, which means there are no gradients, textures, or light sources.</span></span>

## <a name="designing"></a><span data-ttu-id="9d977-130">Проектирование</span><span class="sxs-lookup"><span data-stu-id="9d977-130">Designing</span></span>

### <a name="sizes"></a><span data-ttu-id="9d977-131">Размеры</span><span class="sxs-lookup"><span data-stu-id="9d977-131">Sizes</span></span>

<span data-ttu-id="9d977-132">Мы рекомендуем создавать каждый значок всех этих размеров для поддержки устройств с высоким DPI.</span><span class="sxs-lookup"><span data-stu-id="9d977-132">We recommend that you produce each icon in all these sizes to support high DPI devices.</span></span> <span data-ttu-id="9d977-133">Абсолютно *необходимые* размеры: 16, 20 и 32 пкс, так как это 100% размеров.</span><span class="sxs-lookup"><span data-stu-id="9d977-133">The absolutely *required* sizes are 16 px, 20 px, and 32 px, as those are the 100% sizes.</span></span>

<span data-ttu-id="9d977-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span><span class="sxs-lookup"><span data-stu-id="9d977-134">**16 px, 20 px, 24 px, 32 px, 40 px, 48 px, 64 px, 80 px, 96 px**</span></span>

### <a name="layout"></a><span data-ttu-id="9d977-135">Макет</span><span class="sxs-lookup"><span data-stu-id="9d977-135">Layout</span></span>

<span data-ttu-id="9d977-136">Ниже приводится пример макета значка с модификатором.</span><span class="sxs-lookup"><span data-stu-id="9d977-136">The following is an example of icon layout with a modifier.</span></span>

![Схема значка с модификатором в правом нижнем](../images/monolineicon1.png)  ![Схема того же значка с добавленным фоном сетки и вызовами для базового, модификатора, заполнения и вырезания](../images/monolineicon2.png)

#### <a name="elements"></a><span data-ttu-id="9d977-139">Элементы</span><span class="sxs-lookup"><span data-stu-id="9d977-139">Elements</span></span>

- <span data-ttu-id="9d977-140">**Base**: основная концепция, представленная значком.</span><span class="sxs-lookup"><span data-stu-id="9d977-140">**Base**: The main concept that the icon represents.</span></span> <span data-ttu-id="9d977-141">Обычно это единственный визуальный элемент, необходимый для значка, но иногда основную концепцию можно расширить с помощью дополнительного элемента модификатора.</span><span class="sxs-lookup"><span data-stu-id="9d977-141">This is usually the only visual needed for the icon, but sometimes the main concept can be enhanced with a secondary element, a modifier.</span></span>

- <span data-ttu-id="9d977-142">**Модификатор** Любой элемент, наложение базового; то есть модификатор, который обычно представляет действие или состояние.</span><span class="sxs-lookup"><span data-stu-id="9d977-142">**Modifier** Any element that overlays the base; that is, a modifier that typically represents an action or a status.</span></span> <span data-ttu-id="9d977-143">Он изменяет базовый элемент, выступая в качестве дополнения, изменения или дескриптора.</span><span class="sxs-lookup"><span data-stu-id="9d977-143">It modifies the base element by acting as an addition, alteration, or a descriptor.</span></span>

![Схема сетки с областями базовых и модификаторов](../images/monolineicon3.png)

### <a name="construction"></a><span data-ttu-id="9d977-145">Строительство</span><span class="sxs-lookup"><span data-stu-id="9d977-145">Construction</span></span>

#### <a name="element-placement"></a><span data-ttu-id="9d977-146">Размещение элементов</span><span class="sxs-lookup"><span data-stu-id="9d977-146">Element placement</span></span>

<span data-ttu-id="9d977-147">Базовые элементы размещаются в центре значка в заполнении.</span><span class="sxs-lookup"><span data-stu-id="9d977-147">Base elements are placed in the center of the icon within the padding.</span></span> <span data-ttu-id="9d977-148">Если вы не можете разместить его по центру, база должна переулокаться справа вверху.</span><span class="sxs-lookup"><span data-stu-id="9d977-148">If it can't be placed perfectly centered, then the base should err to the top right.</span></span> <span data-ttu-id="9d977-149">В следующем примере значок идеально по центру.</span><span class="sxs-lookup"><span data-stu-id="9d977-149">In the following example, the icon is perfectly centered.</span></span>

![Схема, на которой показан идеально центрный значок](../images/monolineicon4.png)

<span data-ttu-id="9d977-151">В следующем примере значок находится слева.</span><span class="sxs-lookup"><span data-stu-id="9d977-151">In the following example, the icon is erring to the left.</span></span>

![Схема, на которой показан значок, который перебор слева на 1 пкс](../images/monolineicon5.png)

<span data-ttu-id="9d977-153">Модификаторы почти всегда помещаются в нижний правый угол холста значка.</span><span class="sxs-lookup"><span data-stu-id="9d977-153">Modifiers are almost always placed in the bottom right corner of the icon canvas.</span></span> <span data-ttu-id="9d977-154">В некоторых редких случаях модификаторы помещаются в другой угол.</span><span class="sxs-lookup"><span data-stu-id="9d977-154">In some rare cases, modifiers are placed in a different corner.</span></span> <span data-ttu-id="9d977-155">Например, если базовый элемент будет недостижим с модификатором в правом нижнем углу, рассмотрите возможность размещения его в левом верхнем углу.</span><span class="sxs-lookup"><span data-stu-id="9d977-155">For example, if the base element would be unrecognizable with the modifier in the bottom right corner, then consider placing it in the upper left corner.</span></span>

![Схема, на которой четыре значка с модификатором в правом нижнем и один значок с модификатором в левом верхнем](../images/monolineicon6.png)

#### <a name="padding"></a><span data-ttu-id="9d977-157">Padding</span><span class="sxs-lookup"><span data-stu-id="9d977-157">Padding</span></span>

<span data-ttu-id="9d977-158">Каждый значок размера имеет заданный объем заполнения вокруг значка.</span><span class="sxs-lookup"><span data-stu-id="9d977-158">Each size icon has a specified amount of padding around the icon.</span></span> <span data-ttu-id="9d977-159">Базовый элемент остается в пределах заполнения, но модификатор должен приоставливать его до края холста, расширяя его за пределами отбивки до края границы значка.</span><span class="sxs-lookup"><span data-stu-id="9d977-159">The base element stays within the padding, but the modifier should butt up to the edge of the canvas, extending outside of the padding to the edge of the icon border.</span></span> <span data-ttu-id="9d977-160">На следующих изображениях покажите рекомендуемое заполнение для каждого размера значка.</span><span class="sxs-lookup"><span data-stu-id="9d977-160">The following images show the recommended padding to use for each of the icon sizes.</span></span>

|<span data-ttu-id="9d977-161">**16 пк**</span><span class="sxs-lookup"><span data-stu-id="9d977-161">**16px**</span></span>|<span data-ttu-id="9d977-162">**20 пк**</span><span class="sxs-lookup"><span data-stu-id="9d977-162">**20px**</span></span>|<span data-ttu-id="9d977-163">**24 пк**</span><span class="sxs-lookup"><span data-stu-id="9d977-163">**24px**</span></span>|<span data-ttu-id="9d977-164">**32 пк**</span><span class="sxs-lookup"><span data-stu-id="9d977-164">**32px**</span></span>|<span data-ttu-id="9d977-165">**40 пк**</span><span class="sxs-lookup"><span data-stu-id="9d977-165">**40px**</span></span>|<span data-ttu-id="9d977-166">**48 пк**</span><span class="sxs-lookup"><span data-stu-id="9d977-166">**48px**</span></span>|<span data-ttu-id="9d977-167">**64 пк**</span><span class="sxs-lookup"><span data-stu-id="9d977-167">**64px**</span></span>|<span data-ttu-id="9d977-168">**80 пк**</span><span class="sxs-lookup"><span data-stu-id="9d977-168">**80px**</span></span>|<span data-ttu-id="9d977-169">**96px**</span><span class="sxs-lookup"><span data-stu-id="9d977-169">**96px**</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|![Значок 16 пк с заполнением 0 пк](../images/monolineicon7.png)|![Значок 20 пкс с заполнением 1 пк](../images/monolineicon8.png)|![Значок 24 пкс с заполнением 1 пкс](../images/monolineicon9.png)|![Значок 32 пкс с заполнением 2 пкс](../images/monolineicon10.png)|![Значок 40 пкс с отбивкой 2 пкс](../images/monolineicon11.png)|![Значок 48 пк с заполнением 3 пкс](../images/monolineicon12.png)|![Значок 64 пкс с заполнением 4 пкс](../images/monolineicon13.png)|![Значок 80 пк с заполнением 5 пк](../images/monolineicon14.png)|![Значок 96 пк с заполнением 6 пк](../images/monolineicon15.png)|

#### <a name="line-weights"></a><span data-ttu-id="9d977-179">Вес строки</span><span class="sxs-lookup"><span data-stu-id="9d977-179">Line weights</span></span>

<span data-ttu-id="9d977-180">Monoline — это стиль, в котором фигуры обозначены строками и структурами.</span><span class="sxs-lookup"><span data-stu-id="9d977-180">Monoline is a style dominated by line and outlined shapes.</span></span> <span data-ttu-id="9d977-181">В зависимости от того, какой размер вы производите значок, следует использовать следующие веса строки.</span><span class="sxs-lookup"><span data-stu-id="9d977-181">Depending on what size you are producing the icon should use the following line weights.</span></span>

|<span data-ttu-id="9d977-182">Размер значка:</span><span class="sxs-lookup"><span data-stu-id="9d977-182">Icon Size:</span></span>|<span data-ttu-id="9d977-183">16 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-183">16px</span></span>|<span data-ttu-id="9d977-184">20 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-184">20px</span></span>|<span data-ttu-id="9d977-185">24 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-185">24px</span></span>|<span data-ttu-id="9d977-186">32 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-186">32px</span></span>|<span data-ttu-id="9d977-187">40 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-187">40px</span></span>|<span data-ttu-id="9d977-188">48 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-188">48px</span></span>|<span data-ttu-id="9d977-189">64 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-189">64px</span></span>|<span data-ttu-id="9d977-190">80 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-190">80px</span></span>|<span data-ttu-id="9d977-191">96px</span><span class="sxs-lookup"><span data-stu-id="9d977-191">96px</span></span>|
|:---|:---|:---|:---|:---|:---|:---|:---|:---|:---|
|<span data-ttu-id="9d977-192">**Вес строки:**</span><span class="sxs-lookup"><span data-stu-id="9d977-192">**Line Weight:**</span></span>|<span data-ttu-id="9d977-193">1 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-193">1px</span></span>|<span data-ttu-id="9d977-194">1 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-194">1px</span></span>|<span data-ttu-id="9d977-195">1 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-195">1px</span></span>|<span data-ttu-id="9d977-196">1 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-196">1px</span></span>|<span data-ttu-id="9d977-197">2 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-197">2px</span></span>|<span data-ttu-id="9d977-198">2 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-198">2px</span></span>|<span data-ttu-id="9d977-199">2 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-199">2px</span></span>|<span data-ttu-id="9d977-200">2 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-200">2px</span></span>|<span data-ttu-id="9d977-201">3 пк</span><span class="sxs-lookup"><span data-stu-id="9d977-201">3px</span></span>|
|<span data-ttu-id="9d977-202">**Пример значка:**</span><span class="sxs-lookup"><span data-stu-id="9d977-202">**Example icon:**</span></span>|![Значок 16 пк](../images/monolineicon16.png)|![Значок 20 пкс](../images/monolineicon17.png)|![Значок 24 пк](../images/monolineicon18.png)|![Значок 32 пк](../images/monolineicon19.png)|![Значок 40 пк](../images/monolineicon20.png)|![Значок 48 пк](../images/monolineicon21.png)|![Значок 64 пк](../images/monolineicon22.png)|![Значок 80 пк](../images/monolineicon23.png)|![Значок 96 пк](../images/monolineicon24.png)|

#### <a name="cutouts"></a><span data-ttu-id="9d977-212">Cutouts</span><span class="sxs-lookup"><span data-stu-id="9d977-212">Cutouts</span></span>

<span data-ttu-id="9d977-213">Когда элемент значка помещается поверх другого элемента, вырезание (нижнего элемента) используется для предоставления пространства между двумя элементами, в основном в целях учитаемости.</span><span class="sxs-lookup"><span data-stu-id="9d977-213">When an icon element is placed on top of another element, a cutout (of the bottom element) is used to provide space between the two elements, mainly for readability purposes.</span></span> <span data-ttu-id="9d977-214">Обычно это происходит, когда модификатор помещается поверх базового элемента, но также есть случаи, когда ни один из элементов не является модификатором.</span><span class="sxs-lookup"><span data-stu-id="9d977-214">This usually happens when a modifier is placed on top of a base element, but there are also cases where neither of the elements is a modifier.</span></span> <span data-ttu-id="9d977-215">Эти вырезания между двумя элементами иногда называются разрывом.</span><span class="sxs-lookup"><span data-stu-id="9d977-215">These cutouts between the two elements is sometimes referred to as a "gap".</span></span>

<span data-ttu-id="9d977-216">Размер разрывов должен быть такой же шириной, как и вес линии, используемый для этого размера.</span><span class="sxs-lookup"><span data-stu-id="9d977-216">The size of the gap should be the same width as the line weight used on that size.</span></span> <span data-ttu-id="9d977-217">Если значок составляет 16 пкс, ширина разрывов будет 1 пк, а если это значок 48 пкс, разрыв должен быть 2 пк.</span><span class="sxs-lookup"><span data-stu-id="9d977-217">If making a 16 px icon, the gap width would be 1px and if it is a 48 px icon then the gap should be 2px.</span></span> <span data-ttu-id="9d977-218">В следующем примере показан значок 32 пкс с разрывом в 1 пкс между модификатором и базовой базой.</span><span class="sxs-lookup"><span data-stu-id="9d977-218">The following example shows a 32 px icon with a gap of 1px between the modifier and the underlying base.</span></span>

![Значок 32 пкс с разрывом в 1 пкс между модификатором и базовой базой](../images/monolineicon25.png)

<span data-ttu-id="9d977-220">В некоторых случаях разрыв может увеличиться на 1/2 пкс, если модификатор имеет диагональный или кривый край, а стандартный разрыв не обеспечивает достаточного разделения.</span><span class="sxs-lookup"><span data-stu-id="9d977-220">In some cases, the gap can be increase by a 1/2 px if the modifier has a diagonal or curved edge and the standard gap doesn't provide enough separation.</span></span> <span data-ttu-id="9d977-221">Скорее всего, это повлияет только на значки с весом 1 пкс: 16 пк, 20 пкс, 24 пкс и 32 пкс.</span><span class="sxs-lookup"><span data-stu-id="9d977-221">This will likely only affect the icons with 1px line weight: 16 px, 20 px, 24 px, and 32 px.</span></span>

#### <a name="background-fills"></a><span data-ttu-id="9d977-222">Фоновые заливки</span><span class="sxs-lookup"><span data-stu-id="9d977-222">Background fills</span></span>

<span data-ttu-id="9d977-223">Большинство значков в наборе значков Monoline требуют заполнения фона.</span><span class="sxs-lookup"><span data-stu-id="9d977-223">Most icons in the Monoline icon set require background fills.</span></span> <span data-ttu-id="9d977-224">Однако в некоторых случаях у объекта не будет заливки, поэтому заливка не должна применяться.</span><span class="sxs-lookup"><span data-stu-id="9d977-224">However, there are cases where the object would not naturally have a fill, so no fill should be applied.</span></span> <span data-ttu-id="9d977-225">Следующие значки имеют белый заливки.</span><span class="sxs-lookup"><span data-stu-id="9d977-225">The following icons have a white fill.</span></span>

![Компиляция пяти значков с белыми заливками](../images/monolineicon26.png)

<span data-ttu-id="9d977-227">Следующие значки не заполняются.</span><span class="sxs-lookup"><span data-stu-id="9d977-227">The following icons have no fill.</span></span> <span data-ttu-id="9d977-228">(Значок шестеренки включен, чтобы показать, что центральное изображение не заполнено.)</span><span class="sxs-lookup"><span data-stu-id="9d977-228">(The gear icon is included to show that the center hole is not filled.)</span></span>

![Компиляция пяти значков без заливки](../images/monolineicon27.png)

##### <a name="best-practices-for-fills"></a><span data-ttu-id="9d977-230">Best practices for fills</span><span class="sxs-lookup"><span data-stu-id="9d977-230">Best practices for fills</span></span>

###### <a name="dos"></a><span data-ttu-id="9d977-231">Dos:</span><span class="sxs-lookup"><span data-stu-id="9d977-231">Dos:</span></span>

- <span data-ttu-id="9d977-232">Заполните любой элемент, который имеет запредельную границу и, естественно, будет иметь заливки.</span><span class="sxs-lookup"><span data-stu-id="9d977-232">Fill any element that has a defined boundary, and would naturally have a fill.</span></span>
- <span data-ttu-id="9d977-233">Используйте отдельную фигуру для создания заливки фона.</span><span class="sxs-lookup"><span data-stu-id="9d977-233">Use a separate shape to create the background fill.</span></span>
- <span data-ttu-id="9d977-234">Используйте **фоновую заливку** из [цветовой палитры.](#color)</span><span class="sxs-lookup"><span data-stu-id="9d977-234">Use **Background Fill** from the [color palette](#color).</span></span>
- <span data-ttu-id="9d977-235">Поддерживание разделения пикселей между перекрывающимися элементами.</span><span class="sxs-lookup"><span data-stu-id="9d977-235">Maintain the pixel separation between overlapping elements.</span></span>
- <span data-ttu-id="9d977-236">Заполните несколько объектов.</span><span class="sxs-lookup"><span data-stu-id="9d977-236">Fill between multiple objects.</span></span>

###### <a name="donts"></a><span data-ttu-id="9d977-237">Не:</span><span class="sxs-lookup"><span data-stu-id="9d977-237">Don'ts:</span></span>

- <span data-ttu-id="9d977-238">Не заполняйте объекты, которые не были бы заполнены естественным образом; например, paperclip.</span><span class="sxs-lookup"><span data-stu-id="9d977-238">Don't fill objects that would not naturally be filled; for example, a paperclip.</span></span>
- <span data-ttu-id="9d977-239">Не заполняйте скобки.</span><span class="sxs-lookup"><span data-stu-id="9d977-239">Don't fill brackets.</span></span>
- <span data-ttu-id="9d977-240">Не заполняйте цифры или буквы.</span><span class="sxs-lookup"><span data-stu-id="9d977-240">Don't fill behind numbers or alpha characters.</span></span>

### <a name="color"></a><span data-ttu-id="9d977-241">Цвет</span><span class="sxs-lookup"><span data-stu-id="9d977-241">Color</span></span>

<span data-ttu-id="9d977-242">Цветовая палитра разработана для простоты и доступности.</span><span class="sxs-lookup"><span data-stu-id="9d977-242">The color palette has been designed for simplicity and accessibility.</span></span> <span data-ttu-id="9d977-243">Он содержит 4 нейтральных цвета и два варианта для синего, зеленого, желтого, красного и сиреневых.</span><span class="sxs-lookup"><span data-stu-id="9d977-243">It contains 4 neutral colors and two variations for blue, green, yellow, red, and purple.</span></span> <span data-ttu-id="9d977-244">Оранжевый цвет намеренно не включен в цветовую палитру монолайновых значков.</span><span class="sxs-lookup"><span data-stu-id="9d977-244">Orange is intentionally not included in the Monoline icon color palette.</span></span> <span data-ttu-id="9d977-245">Каждый цвет предназначен для использования определенными способами, как описано в этом разделе.</span><span class="sxs-lookup"><span data-stu-id="9d977-245">Each color is intended to be used in specific ways as outlined in this section.</span></span>

#### <a name="palette"></a><span data-ttu-id="9d977-246">Палитра</span><span class="sxs-lookup"><span data-stu-id="9d977-246">Palette</span></span>

![Четыре оттенка серого в монолайне: темно-серый для автономных или контурных, средний серый для контура или содержимого, очень светло-серый для заливки фона и светло-серый для заливки](../images/monoline-grayshades.png)

![Цветовая палитра в монолайне включает в себя синий, зеленый, желтый, красный и сиреневый цвет для автономных, контурных и заливки](../images/monoline-colors.png)

#### <a name="how-to-use-color"></a><span data-ttu-id="9d977-249">Использование цвета</span><span class="sxs-lookup"><span data-stu-id="9d977-249">How to use color</span></span>

<span data-ttu-id="9d977-250">В монолайновой цветовой палитре все цвета имеют автономные варианты, варианты Outline и Fill.</span><span class="sxs-lookup"><span data-stu-id="9d977-250">In the Monoline color palette, all colors have Standalone, Outline, and Fill variations.</span></span> <span data-ttu-id="9d977-251">Как правило, элементы построены с помощью заливки и границы.</span><span class="sxs-lookup"><span data-stu-id="9d977-251">Generally, elements are constructed with a fill and a border.</span></span> <span data-ttu-id="9d977-252">Цвета применяются в одном из следующих шаблонов:</span><span class="sxs-lookup"><span data-stu-id="9d977-252">The colors are applied in one of the following patterns:</span></span>

- <span data-ttu-id="9d977-253">Автономный цвет только для объектов без заливки.</span><span class="sxs-lookup"><span data-stu-id="9d977-253">The Standalone color alone for objects that have no fill.</span></span>
- <span data-ttu-id="9d977-254">На границе используется цвет Outline, а для заливки используется цвет заливки.</span><span class="sxs-lookup"><span data-stu-id="9d977-254">The border uses the Outline color and the fill uses the Fill color.</span></span>
- <span data-ttu-id="9d977-255">На границе используется автономный цвет, а для заливки используется цвет заливки фона.</span><span class="sxs-lookup"><span data-stu-id="9d977-255">The border uses the Standalone color and the fill uses the Background Fill color.</span></span>

<span data-ttu-id="9d977-256">Ниже приводится пример использования цвета.</span><span class="sxs-lookup"><span data-stu-id="9d977-256">The following are examples of using color.</span></span>

![Компиляция трех значков с цветом на границе или заливке или и тем, и другим](../images/monolineicon28.png)

<span data-ttu-id="9d977-258">Наиболее распространенной ситуацией является использование элементом standalone темно-серого цвета с фоновой заливки.</span><span class="sxs-lookup"><span data-stu-id="9d977-258">The most common situation will be to have an element use Dark Gray Standalone with Background Fill.</span></span>

<span data-ttu-id="9d977-259">При использовании цветного заливки он всегда должен быть с соответствующим цветом Outline.</span><span class="sxs-lookup"><span data-stu-id="9d977-259">When using a colored Fill, it should always be with its corresponding Outline color.</span></span> <span data-ttu-id="9d977-260">Например, синяя заливка должна использоваться только с синим контуром.</span><span class="sxs-lookup"><span data-stu-id="9d977-260">For example, Blue Fill should only be used with Blue Outline.</span></span> <span data-ttu-id="9d977-261">Однако есть два исключения из этого общего правила:</span><span class="sxs-lookup"><span data-stu-id="9d977-261">But there are two exceptions to this general rule:</span></span>

- <span data-ttu-id="9d977-262">Фоновое заполнение можно использовать с любым автономным цветом.</span><span class="sxs-lookup"><span data-stu-id="9d977-262">Background Fill can be used with any color Standalone.</span></span>
- <span data-ttu-id="9d977-263">Светло-серую заливка можно использовать с двумя разными цветами outline: темно-серым или средним серым.</span><span class="sxs-lookup"><span data-stu-id="9d977-263">Light Gray Fill can be used with two different Outline colors: Dark Gray or Medium Gray.</span></span>

#### <a name="when-to-use-color"></a><span data-ttu-id="9d977-264">Когда использовать цвет</span><span class="sxs-lookup"><span data-stu-id="9d977-264">When to use color</span></span>

<span data-ttu-id="9d977-265">Цвет следует использовать для передачи значения значка, а не для замеления.</span><span class="sxs-lookup"><span data-stu-id="9d977-265">Color should be used to convey the meaning of the icon rather than for embellishment.</span></span> <span data-ttu-id="9d977-266">Он должен **выделить действие** для пользователя.</span><span class="sxs-lookup"><span data-stu-id="9d977-266">It should **highlight the action** to the user.</span></span> <span data-ttu-id="9d977-267">При добавлении модификатора в базовый элемент с цветом базовый элемент обычно превращается в темно-серый и фоновый заливка, чтобы модификатором был элемент цвета, например в приведенном ниже примере с модификатором "X", добавляемого в базу рисунков в левом значке следующего набора.</span><span class="sxs-lookup"><span data-stu-id="9d977-267">When a modifier is added to a base element that has color, the base element is typically turned into Dark Gray and Background Fill so that the modifier can be the element of color, such as the case below with the "X" modifier being added to the picture base in the leftmost icon of the following set.</span></span>

![Компиляция пяти значков, которые используют цвет](../images/monolineicon29.png)

<span data-ttu-id="9d977-269">Значки следует ограничить одним **дополнительным** цветом, кроме упомянутых выше окрашиваний Outline и Fill.</span><span class="sxs-lookup"><span data-stu-id="9d977-269">You should limit your icons to **one** additional color, other than the Outline and Fill mentioned above.</span></span> <span data-ttu-id="9d977-270">Тем не менее, можно использовать больше цветов, если это важно для его метафоры, с ограничением двух дополнительных цветов, кроме серого.</span><span class="sxs-lookup"><span data-stu-id="9d977-270">However, more colors can be used if it is vital for its metaphor, with a limit of two additional colors other than gray.</span></span> <span data-ttu-id="9d977-271">В редких случаях существуют исключения, когда требуется больше цветов.</span><span class="sxs-lookup"><span data-stu-id="9d977-271">In rare cases, there are exceptions when more colors are needed.</span></span> <span data-ttu-id="9d977-272">Ниже приводится хороший пример значков, которые используют только один цвет.</span><span class="sxs-lookup"><span data-stu-id="9d977-272">The following are good examples of icons that use just one color.</span></span>

  ![Компиляция пяти значков, каждый из которых использует один цвет](../images/monolineicon30.png)

<span data-ttu-id="9d977-274">Но в следующих значках используется слишком много цветов.</span><span class="sxs-lookup"><span data-stu-id="9d977-274">But the following icons use too many colors.</span></span>

  ![Компиляция пяти значков, каждый из которых использует несколько цветов](../images/monolineicon31.png)

<span data-ttu-id="9d977-276">Используйте **средний серый** цвет для внутреннего "содержимого", например линий сетки в значке таблицы.</span><span class="sxs-lookup"><span data-stu-id="9d977-276">Use **Medium Gray** for interior "content", such as grid lines in an icon of a spreadsheet.</span></span> <span data-ttu-id="9d977-277">Дополнительные внутренние цвета используются, когда содержимое должно показывать поведение этого средства управления.</span><span class="sxs-lookup"><span data-stu-id="9d977-277">Additional interior colors are used when the content needs to show the behavior of the control.</span></span>

![Компиляция пяти значков со средним серым цветом внутренних элементов](../images/monolineicon32.png)

#### <a name="text-lines"></a><span data-ttu-id="9d977-279">Текстовые строки</span><span class="sxs-lookup"><span data-stu-id="9d977-279">Text lines</span></span>

<span data-ttu-id="9d977-280">Если текстовые строки находятся в "контейнере" (например, текст в документе), используйте средний серый цвет.</span><span class="sxs-lookup"><span data-stu-id="9d977-280">When text lines are in a "container" (for example, text on a document), use medium gray.</span></span> <span data-ttu-id="9d977-281">Текстовые строки, не в контейнере, должны быть **темно-серыми.**</span><span class="sxs-lookup"><span data-stu-id="9d977-281">Text lines not in a container should be **Dark Gray**.</span></span>

### <a name="text"></a><span data-ttu-id="9d977-282">Текст</span><span class="sxs-lookup"><span data-stu-id="9d977-282">Text</span></span>

<span data-ttu-id="9d977-283">Избегайте использования текстовых символов в значках.</span><span class="sxs-lookup"><span data-stu-id="9d977-283">Avoid using text characters in icons.</span></span> <span data-ttu-id="9d977-284">Так как продукты Office используются по всему миру, мы хотим, чтобы значки были максимально нейтральными на языке.</span><span class="sxs-lookup"><span data-stu-id="9d977-284">Since Office products are used around the world, we want to keep icons as language neutral as possible.</span></span>

## <a name="production"></a><span data-ttu-id="9d977-285">Производство</span><span class="sxs-lookup"><span data-stu-id="9d977-285">Production</span></span>

### <a name="icon-file-format"></a><span data-ttu-id="9d977-286">Формат файла значка</span><span class="sxs-lookup"><span data-stu-id="9d977-286">Icon file format</span></span>

<span data-ttu-id="9d977-287">Конечные значки должны быть сохранены в PNG-файлах изображений.</span><span class="sxs-lookup"><span data-stu-id="9d977-287">The final icons should be saved as .png image files.</span></span> <span data-ttu-id="9d977-288">Используйте формат PNG с прозрачным фоном и 32-битной глубиной.</span><span class="sxs-lookup"><span data-stu-id="9d977-288">Use PNG format with a transparent background and have 32-bit depth.</span></span>
