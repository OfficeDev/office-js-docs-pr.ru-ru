---
title: Использование движения в надстройках Office
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d3be2454b36fe1003c0697f0bca3c29d743e5330
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449063"
---
# <a name="using-motion-in-office-add-ins"></a><span data-ttu-id="0ecac-102">Использование движения в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="0ecac-102">Using motion in Office Add-ins</span></span>

<span data-ttu-id="0ecac-p101">Вы можете использовать движение, чтобы сделать надстройку Office удобнее для пользователя. Элементы пользовательского интерфейса, элементы управления и компоненты часто отличаются интерактивным поведением, требующим переходов, перемещений или анимации. Общие характеристики перемещения между элементами пользовательского интерфейса определяют свойства анимации языка дизайна.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p101">When you design an Office Add-in, you can use motion to enhance the user experience. UI elements, controls, and components often have interactive behaviors that require transitions, motion, or animation. Common characteristics of motion across UI elements define the animation aspects of a design language.</span></span> 

<span data-ttu-id="0ecac-p102">Так как набор Office ориентирован на производительность, язык анимации Office нацелен в первую очередь на выполнение клиентами своих задач. Он обеспечивает баланс между оперативным откликом, надежной хореографией и удобством использования. Внедренные в Office надстройки работают в контексте этого языка анимации. Поэтому, применяя движение, важно учитывать указанные ниже рекомендации.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p102">Because Office is focused on productivity, the Office animation language supports the goal of helping customers get things done. It strikes a balance between performant response, reliable choreography, and detailed delight. Add-ins embedded in Office sit within this existing animation language. Given this context, it is important to consider the following guidelines when applying motion.</span></span> 


## <a name="create-motion-with-a-purpose"></a><span data-ttu-id="0ecac-110">Создавайте движение с определенной целью</span><span class="sxs-lookup"><span data-stu-id="0ecac-110">Create motion with a purpose</span></span>

<span data-ttu-id="0ecac-p103">Движение должно иметь цель, представляющую ценность для пользователя. Учитывайте тон и цель содержимого при выборе анимации. Обрабатывайте критические сообщения не так, как описательные.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p103">Motion should have a purpose that communicates additional value to the user. Consider the tone and purpose of your content when choosing animations. Handle critical messages differently than exploratory navigations.</span></span>

<span data-ttu-id="0ecac-p104">Стандартные элементы, используемые в надстройке, могут включать движение, которое акцентирует внимание пользователя, показывает, как элементы связаны друг с другом, или подтверждает правильность действия. Спланируйте хореографию элементов, чтобы усилить иерархию и умозрительные модели.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p104">Standard elements used in an add-in can incorporate motion to help focus the user, show how elements relate to each other, and validate user actions. Choreograph elements to reinforce hierarchy and mental models.</span></span>

### <a name="best-practices"></a><span data-ttu-id="0ecac-116">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="0ecac-116">Best practices</span></span>

|<span data-ttu-id="0ecac-117">Правильно</span><span class="sxs-lookup"><span data-stu-id="0ecac-117">Do</span></span>|<span data-ttu-id="0ecac-118">Неправильно</span><span class="sxs-lookup"><span data-stu-id="0ecac-118">Don't</span></span>|
|:-----|:-----|
|<span data-ttu-id="0ecac-p105">Определите основные элементы надстройки, которые нужно анимировать. Обычно анимируются панели, оверлеи, модальные окна, подсказки, меню и учебные выноски.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p105">Identify key elements in the add-in that should have motion. Commonly animated elements in an add-in are panels, overlays, modals, tool tips, menus, and teaching call outs.</span></span>| <span data-ttu-id="0ecac-p106">Не перегружайте пользователя, анимируя все элементы. Не применяйте нескольких движений, которые акцентируют внимание пользователя на нескольких элементах одновременно.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p106">Don't overwhelm the user by animating every element. Avoid applying multiple motions that attempt to lead or focus the user on many elements at once.</span></span> |
|<span data-ttu-id="0ecac-p107">Используйте простое предсказуемое движение. Учитывайте происхождение элемента-триггера. Используйте движение, чтобы создать связь между действием и итоговым пользовательским интерфейсом.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p107">Use simple, subtle motion that behaves in expected ways. Consider the origin of your triggering element. Use motion to create a link between the action and the resulting UI.</span></span> | <span data-ttu-id="0ecac-p108">Не заставляйте пользователя ждать движения. Движение в надстройках не должно препятствовать выполнению задачи.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p108">Don't create wait time for a motion. Motion in add-ins should not hinder task completion.</span></span>|

![Открытая панель с минимальным количеством движущихся элементов рядом с открытой панелью с большим количеством движущихся элементов](../images/add-in-motion-purpose.gif)

## <a name="use-expected-motions"></a><span data-ttu-id="0ecac-129">Используйте предсказуемые движения</span><span class="sxs-lookup"><span data-stu-id="0ecac-129">Use expected motions</span></span>

<span data-ttu-id="0ecac-130">Рекомендуем использовать [Office UI Fabric](https://developer.microsoft.com/fabric) для создания визуальной связи с платформой Office, а также [анимации Fabric](https://developer.microsoft.com/fabric#/styles/animations) для создания движений, которые согласуются с языком движения Fabric.</span><span class="sxs-lookup"><span data-stu-id="0ecac-130">We recommend using [Office UI Fabric](https://developer.microsoft.com/fabric) to create a visual connection with the Office platform, and we also encourage the use of [Fabric Animations](https://developer.microsoft.com/fabric#/styles/animations) to create motions that align with the Fabric motion language.</span></span> 

<span data-ttu-id="0ecac-p109">Используйте эту платформу для более простой интеграции с Office. Это поможет создавать удобные в работе интерфейсы. Классы CSS анимации обеспечивают направленность, точки входа и выхода, а также особенности длительности, которые усиливают умозрительные модели Office и помогают пользователям научиться работать с вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p109">Use it to fit seamlessly in Office. It will help you create experiences that are more felt than observed. The animation CSS classes provide directionality, enter/exit, and duration specifics that reinforce Office mental models and provide opportunities for customers to learn how to interact with your add-in.</span></span>

### <a name="best-practices"></a><span data-ttu-id="0ecac-134">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="0ecac-134">Best practices</span></span>

|<span data-ttu-id="0ecac-135">Правильно</span><span class="sxs-lookup"><span data-stu-id="0ecac-135">Do</span></span>|<span data-ttu-id="0ecac-136">Неправильно</span><span class="sxs-lookup"><span data-stu-id="0ecac-136">Don't</span></span>|
|:-----|:-----|
|<span data-ttu-id="0ecac-137">Используйте движение, которое согласуется с языком движения Fabric.</span><span class="sxs-lookup"><span data-stu-id="0ecac-137">Use motion that aligns with behaviors in Fabric.</span></span>| <span data-ttu-id="0ecac-138">Не создавайте движения, которые конфликтуют со стандартными шаблонами движения в Office.</span><span class="sxs-lookup"><span data-stu-id="0ecac-138">Don't create motions that interfere or conflict with common motion patterns in Office.</span></span>
|<span data-ttu-id="0ecac-139">Убедитесь, что существует согласованное приложение движения между элементами Like.</span><span class="sxs-lookup"><span data-stu-id="0ecac-139">Ensure that there is a consistent application of motion across like elements.</span></span>| <span data-ttu-id="0ecac-140">Не используйте разные движения для анимации одного и того же компонента или объекта.</span><span class="sxs-lookup"><span data-stu-id="0ecac-140">Don't use different motions to animate the same component or object.</span></span>|
|<span data-ttu-id="0ecac-p110">Используйте одно направление при анимации элемента. Например, панель, которая открывается справа, должна закрываться справа.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p110">Create consistency with use of direction in animation. For example, a panel that opens from the right should close to the right.</span></span>|<span data-ttu-id="0ecac-143">Не анимируйте элемент, используя несколько направлений.</span><span class="sxs-lookup"><span data-stu-id="0ecac-143">Don't animate an element using multiple directions.</span></span>

![Предсказуемое и непредсказуемое открытие модального окна](../images/add-in-motion-expected.gif)

## <a name="avoid-out-of-character-motion-for-an-element"></a><span data-ttu-id="0ecac-145">Не используйте движение, которое нетипично для элемента</span><span class="sxs-lookup"><span data-stu-id="0ecac-145">Avoid out of character motion for an element</span></span>

<span data-ttu-id="0ecac-p111">Анимируя элемент, учитывайте размер холста HTML (панели задач, диалогового окна или контентной надстройки). Не перегружайте холст. Движущиеся элементы должны сочетаться со средой Office. Характер движения надстройки должен быть эффективным, надежным и плавным. Стремитесь информировать и направлять пользователя, не осложняя его работу.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p111">Consider the size of the HTML canvas (task pane, dialog box, or content add-in) when implementing motion. Avoid overloading in constrained spaces. Moving element(s) should be in tune with Office. The character of add-in motion should be performant, reliable, and fluid. Instead of impeding productivity, aim to inform and direct.</span></span>

### <a name="best-practices"></a><span data-ttu-id="0ecac-151">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="0ecac-151">Best practices</span></span>

|<span data-ttu-id="0ecac-152">Правильно</span><span class="sxs-lookup"><span data-stu-id="0ecac-152">Do</span></span>|<span data-ttu-id="0ecac-153">Неправильно</span><span class="sxs-lookup"><span data-stu-id="0ecac-153">Don't</span></span>|
|:-----|:-----|
| <span data-ttu-id="0ecac-154">Используйте [рекомендуемую длительность движения](https://developer.microsoft.com/fabric#/styles/animations).</span><span class="sxs-lookup"><span data-stu-id="0ecac-154">Use [recommended motion durations](https://developer.microsoft.com/fabric#/styles/animations).</span></span> | <span data-ttu-id="0ecac-p112">Не используйте чрезмерную анимацию. Старайтесь не создавать нефункциональные движения, которые только отвлекают пользователей.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p112">Don't use exaggerated animations. Avoid creating experiences that embellish and distract your customers.</span></span>
| <span data-ttu-id="0ecac-157">Используйте [рекомендуемые кривые замедления](/windows/uwp/design/motion/timing-and-easing#easing-in-fluent-motion).</span><span class="sxs-lookup"><span data-stu-id="0ecac-157">Follow [recommended easing curves](/windows/uwp/design/motion/timing-and-easing#easing-in-fluent-motion).</span></span>  |<span data-ttu-id="0ecac-p113">Не перемещайте элементы рывками или по частям. Избегайте упреждения, возвратов, эффекта "резиновой ленты" или других эффектов, которые имитируют законы физики реального мира.</span><span class="sxs-lookup"><span data-stu-id="0ecac-p113">Don't move elements in a jerky or disjointed manner. Avoid anticipations, bounces, rubberband, or other effects that emulate natural world physics.</span></span>|

![Загрузка плиток с мягким затуханием и загрузка плиток с отскоком](../images/add-in-motion-character.gif)

## <a name="see-also"></a><span data-ttu-id="0ecac-161">См. также</span><span class="sxs-lookup"><span data-stu-id="0ecac-161">See also</span></span>

* [<span data-ttu-id="0ecac-162">Правила анимации Fabric</span><span class="sxs-lookup"><span data-stu-id="0ecac-162">Fabric animation guidelines</span></span>](https://developer.microsoft.com/fabric#/styles/animations)
* [<span data-ttu-id="0ecac-163">Движение для приложений универсальной платформы Windows</span><span class="sxs-lookup"><span data-stu-id="0ecac-163">Motion for Universal Windows Platform apps</span></span>](/windows/uwp/design/motion)
