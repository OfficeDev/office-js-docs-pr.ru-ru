---
title: Использование движения в надстройках Office
description: Ознакомьтесь с рекомендациями по использованию переходов, движений или анимации в надстройках Office.
ms.date: 07/19/2019
localization_priority: Normal
ms.openlocfilehash: effbd2c3e1e811d9f73c345c80a55a350db909d2
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719268"
---
# <a name="using-motion-in-office-add-ins"></a><span data-ttu-id="14ee7-103">Использование движения в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="14ee7-103">Using motion in Office Add-ins</span></span>

<span data-ttu-id="14ee7-p101">Вы можете использовать движение, чтобы сделать надстройку Office удобнее для пользователя. Элементы пользовательского интерфейса, элементы управления и компоненты часто отличаются интерактивным поведением, требующим переходов, перемещений или анимации. Общие характеристики перемещения между элементами пользовательского интерфейса определяют свойства анимации языка дизайна.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p101">When you design an Office Add-in, you can use motion to enhance the user experience. UI elements, controls, and components often have interactive behaviors that require transitions, motion, or animation. Common characteristics of motion across UI elements define the animation aspects of a design language.</span></span>

<span data-ttu-id="14ee7-p102">Так как набор Office ориентирован на производительность, язык анимации Office нацелен в первую очередь на выполнение клиентами своих задач. Он обеспечивает баланс между оперативным откликом, надежной хореографией и удобством использования. Внедренные в Office надстройки работают в контексте этого языка анимации. Поэтому, применяя движение, важно учитывать указанные ниже рекомендации.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p102">Because Office is focused on productivity, the Office animation language supports the goal of helping customers get things done. It strikes a balance between performant response, reliable choreography, and detailed delight. Add-ins embedded in Office sit within this existing animation language. Given this context, it is important to consider the following guidelines when applying motion.</span></span>


## <a name="create-motion-with-a-purpose"></a><span data-ttu-id="14ee7-111">Создавайте движение с определенной целью</span><span class="sxs-lookup"><span data-stu-id="14ee7-111">Create motion with a purpose</span></span>

<span data-ttu-id="14ee7-p103">Движение должно иметь цель, представляющую ценность для пользователя. Учитывайте тон и цель содержимого при выборе анимации. Обрабатывайте критические сообщения не так, как описательные.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p103">Motion should have a purpose that communicates additional value to the user. Consider the tone and purpose of your content when choosing animations. Handle critical messages differently than exploratory navigations.</span></span>

<span data-ttu-id="14ee7-p104">Стандартные элементы, используемые в надстройке, могут включать движение, которое акцентирует внимание пользователя, показывает, как элементы связаны друг с другом, или подтверждает правильность действия. Спланируйте хореографию элементов, чтобы усилить иерархию и умозрительные модели.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p104">Standard elements used in an add-in can incorporate motion to help focus the user, show how elements relate to each other, and validate user actions. Choreograph elements to reinforce hierarchy and mental models.</span></span>

### <a name="best-practices"></a><span data-ttu-id="14ee7-117">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="14ee7-117">Best practices</span></span>

|<span data-ttu-id="14ee7-118">Правильно</span><span class="sxs-lookup"><span data-stu-id="14ee7-118">Do</span></span>|<span data-ttu-id="14ee7-119">Неправильно</span><span class="sxs-lookup"><span data-stu-id="14ee7-119">Don't</span></span>|
|:-----|:-----|
|<span data-ttu-id="14ee7-p105">Определите основные элементы надстройки, которые нужно анимировать. Обычно анимируются панели, оверлеи, модальные окна, подсказки, меню и учебные выноски.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p105">Identify key elements in the add-in that should have motion. Commonly animated elements in an add-in are panels, overlays, modals, tool tips, menus, and teaching call outs.</span></span>| <span data-ttu-id="14ee7-p106">Не перегружайте пользователя, анимируя все элементы. Не применяйте нескольких движений, которые акцентируют внимание пользователя на нескольких элементах одновременно.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p106">Don't overwhelm the user by animating every element. Avoid applying multiple motions that attempt to lead or focus the user on many elements at once.</span></span> |
|<span data-ttu-id="14ee7-p107">Используйте простое предсказуемое движение. Учитывайте происхождение элемента-триггера. Используйте движение, чтобы создать связь между действием и итоговым пользовательским интерфейсом.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p107">Use simple, subtle motion that behaves in expected ways. Consider the origin of your triggering element. Use motion to create a link between the action and the resulting UI.</span></span> | <span data-ttu-id="14ee7-p108">Не заставляйте пользователя ждать движения. Движение в надстройках не должно препятствовать выполнению задачи.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p108">Don't create wait time for a motion. Motion in add-ins should not hinder task completion.</span></span>|

![Открытая панель с минимальным количеством движущихся элементов рядом с открытой панелью с большим количеством движущихся элементов](../images/add-in-motion-purpose.gif)

## <a name="use-expected-motions"></a><span data-ttu-id="14ee7-130">Используйте предсказуемые движения</span><span class="sxs-lookup"><span data-stu-id="14ee7-130">Use expected motions</span></span>

<span data-ttu-id="14ee7-131">Рекомендуем использовать [Office UI Fabric](https://developer.microsoft.com/fabric) для создания визуальной связи с платформой Office, а также [анимации Fabric](https://developer.microsoft.com/fabric#/styles/web/motion) для создания движений, которые согласуются с языком движения Fabric.</span><span class="sxs-lookup"><span data-stu-id="14ee7-131">We recommend using [Office UI Fabric](https://developer.microsoft.com/fabric) to create a visual connection with the Office platform, and we also encourage the use of [Fabric Animations](https://developer.microsoft.com/fabric#/styles/web/motion) to create motions that align with the Fabric motion language.</span></span>

<span data-ttu-id="14ee7-p109">Используйте эту платформу для более простой интеграции с Office. Это поможет создавать удобные в работе интерфейсы. Классы CSS анимации обеспечивают направленность, точки входа и выхода, а также особенности длительности, которые усиливают умозрительные модели Office и помогают пользователям научиться работать с вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p109">Use it to fit seamlessly in Office. It will help you create experiences that are more felt than observed. The animation CSS classes provide directionality, enter/exit, and duration specifics that reinforce Office mental models and provide opportunities for customers to learn how to interact with your add-in.</span></span>

### <a name="best-practices"></a><span data-ttu-id="14ee7-135">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="14ee7-135">Best practices</span></span>

|<span data-ttu-id="14ee7-136">Правильно</span><span class="sxs-lookup"><span data-stu-id="14ee7-136">Do</span></span>|<span data-ttu-id="14ee7-137">Неправильно</span><span class="sxs-lookup"><span data-stu-id="14ee7-137">Don't</span></span>|
|:-----|:-----|
|<span data-ttu-id="14ee7-138">Используйте движение, которое согласуется с языком движения Fabric.</span><span class="sxs-lookup"><span data-stu-id="14ee7-138">Use motion that aligns with behaviors in Fabric.</span></span>| <span data-ttu-id="14ee7-139">Не создавайте движения, которые конфликтуют со стандартными шаблонами движения в Office.</span><span class="sxs-lookup"><span data-stu-id="14ee7-139">Don't create motions that interfere or conflict with common motion patterns in Office.</span></span>
|<span data-ttu-id="14ee7-140">Убедитесь, что существует согласованное приложение движения между элементами Like.</span><span class="sxs-lookup"><span data-stu-id="14ee7-140">Ensure that there is a consistent application of motion across like elements.</span></span>| <span data-ttu-id="14ee7-141">Не используйте разные движения для анимации одного и того же компонента или объекта.</span><span class="sxs-lookup"><span data-stu-id="14ee7-141">Don't use different motions to animate the same component or object.</span></span>|
|<span data-ttu-id="14ee7-p110">Используйте одно направление при анимации элемента. Например, панель, которая открывается справа, должна закрываться справа.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p110">Create consistency with use of direction in animation. For example, a panel that opens from the right should close to the right.</span></span>|<span data-ttu-id="14ee7-144">Не анимируйте элемент, используя несколько направлений.</span><span class="sxs-lookup"><span data-stu-id="14ee7-144">Don't animate an element using multiple directions.</span></span>

![Предсказуемое и непредсказуемое открытие модального окна](../images/add-in-motion-expected.gif)

## <a name="avoid-out-of-character-motion-for-an-element"></a><span data-ttu-id="14ee7-146">Не используйте движение, которое нетипично для элемента</span><span class="sxs-lookup"><span data-stu-id="14ee7-146">Avoid out of character motion for an element</span></span>

<span data-ttu-id="14ee7-p111">Анимируя элемент, учитывайте размер холста HTML (панели задач, диалогового окна или контентной надстройки). Не перегружайте холст. Движущиеся элементы должны сочетаться со средой Office. Характер движения надстройки должен быть эффективным, надежным и плавным. Стремитесь информировать и направлять пользователя, не осложняя его работу.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p111">Consider the size of the HTML canvas (task pane, dialog box, or content add-in) when implementing motion. Avoid overloading in constrained spaces. Moving element(s) should be in tune with Office. The character of add-in motion should be performant, reliable, and fluid. Instead of impeding productivity, aim to inform and direct.</span></span>

### <a name="best-practices"></a><span data-ttu-id="14ee7-152">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="14ee7-152">Best practices</span></span>

|<span data-ttu-id="14ee7-153">Правильно</span><span class="sxs-lookup"><span data-stu-id="14ee7-153">Do</span></span>|<span data-ttu-id="14ee7-154">Неправильно</span><span class="sxs-lookup"><span data-stu-id="14ee7-154">Don't</span></span>|
|:-----|:-----|
| <span data-ttu-id="14ee7-155">Используйте [рекомендуемую длительность движения](https://developer.microsoft.com/fabric#/styles/web/motion).</span><span class="sxs-lookup"><span data-stu-id="14ee7-155">Use [recommended motion durations](https://developer.microsoft.com/fabric#/styles/web/motion).</span></span> | <span data-ttu-id="14ee7-p112">Не используйте чрезмерную анимацию. Старайтесь не создавать нефункциональные движения, которые только отвлекают пользователей.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p112">Don't use exaggerated animations. Avoid creating experiences that embellish and distract your customers.</span></span>
| <span data-ttu-id="14ee7-158">Используйте [рекомендуемые кривые замедления](/windows/uwp/design/motion/timing-and-easing#easing-in-fluent-motion).</span><span class="sxs-lookup"><span data-stu-id="14ee7-158">Follow [recommended easing curves](/windows/uwp/design/motion/timing-and-easing#easing-in-fluent-motion).</span></span>  |<span data-ttu-id="14ee7-p113">Не перемещайте элементы рывками или по частям. Избегайте упреждения, возвратов, эффекта "резиновой ленты" или других эффектов, которые имитируют законы физики реального мира.</span><span class="sxs-lookup"><span data-stu-id="14ee7-p113">Don't move elements in a jerky or disjointed manner. Avoid anticipations, bounces, rubberband, or other effects that emulate natural world physics.</span></span>|

![Загрузка плиток с мягким затуханием и загрузка плиток с отскоком](../images/add-in-motion-character.gif)

## <a name="see-also"></a><span data-ttu-id="14ee7-162">См. также</span><span class="sxs-lookup"><span data-stu-id="14ee7-162">See also</span></span>

* [<span data-ttu-id="14ee7-163">Правила анимации Fabric</span><span class="sxs-lookup"><span data-stu-id="14ee7-163">Fabric animation guidelines</span></span>](https://developer.microsoft.com/fabric#/styles/web/motion)
* [<span data-ttu-id="14ee7-164">Движение для приложений универсальной платформы Windows</span><span class="sxs-lookup"><span data-stu-id="14ee7-164">Motion for Universal Windows Platform apps</span></span>](/windows/uwp/design/motion)
