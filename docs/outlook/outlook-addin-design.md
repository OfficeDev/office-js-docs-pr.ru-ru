---
title: Дизайн надстройки Outlook
description: Эти рекомендации помогут вам спроектировать и выполнить сборку надстройки с привлекательным внешним видом, которая сделает лучшие качества вашего приложения доступными в Outlook для Windows, веб, iOS, Mac и Android.
ms.date: 06/24/2019
localization_priority: Priority
ms.openlocfilehash: efedeb32643bff12e167931ac4da80fdcc2c277f
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166666"
---
# <a name="outlook-add-in-design-guidelines"></a><span data-ttu-id="776a2-103">Рекомендации по разработке надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="776a2-103">Outlook add-in design guidelines</span></span>

<span data-ttu-id="776a2-p101">Надстройки — отличный способ дополнения базового набора функций Outlook. С помощью надстроек пользователи могут получать доступ к интерфейсам, задачам и содержимому от сторонних разработчиков, не покидая папку "Входящие". После установки надстройки Outlook становятся доступны на всех платформах и устройствах.</span><span class="sxs-lookup"><span data-stu-id="776a2-p101">Add-ins are a great way for partners to extend the functionality of Outlook beyond our core feature set. Add-ins enable users to access third-party experiences, tasks, and content without needing to leave their inbox. Once installed, Outlook add-ins are available on every platform and device.</span></span>  

<span data-ttu-id="776a2-107">Представленные ниже общие рекомендации помогут вам спроектировать и создать привлекательную надстройку, переносящую лучшие аспекты вашего приложения непосредственно в Outlook&mdash;в Windows, Интернете, iOS, Mac и Android.</span><span class="sxs-lookup"><span data-stu-id="776a2-107">The following high-level guidelines will help you design and build a compelling add-in, which brings the best of your app right into Outlook&mdash;on Windows, Web, iOS, Mac, and Android.</span></span>

## <a name="principles"></a><span data-ttu-id="776a2-108">Принципы</span><span class="sxs-lookup"><span data-stu-id="776a2-108">Principles</span></span>

1. <span data-ttu-id="776a2-109">**Сосредоточьтесь на нескольких ключевых задачах, выполняйте их хорошо**</span><span class="sxs-lookup"><span data-stu-id="776a2-109">**Focus on a few key tasks; do them well**</span></span>

   <span data-ttu-id="776a2-p102">Лучшие надстройки просты в использовании, спроектированы с учетом определенных целей и имеют ценность для пользователей. Так как надстройка будет работать в Outlook, этому принципу следует уделить особое внимание. Outlook&mdash;приложение для эффективной работы. Открывая его, пользователи хотят добиться результатов.</span><span class="sxs-lookup"><span data-stu-id="776a2-p102">The best designed add-ins are simple to use, focused, and provide real value to users. Because your add-in will run inside of Outlook, there is additional emphasis placed on this principle. Outlook is a productivity app&mdash;it's where people go to get things done.</span></span>

   <span data-ttu-id="776a2-p103">Ваша надстройка будет расширением нашего интерфейса, поэтому важно убедиться, что предоставляемые ею возможности органично вписываются в Outlook. Задумайтесь, какие из распространенных вариантов использования будет полезнее всего связать с интерфейсами электронной почты и календарей.</span><span class="sxs-lookup"><span data-stu-id="776a2-p103">You will be an extension of our experience and it is important to make sure the scenarios you enable feel like a natural fit inside of Outlook. Think carefully about which of your common use cases will benefit the most from having hooks to them from within our email and calendaring experiences.</span></span>

   <span data-ttu-id="776a2-p104">Не обязательно включать в надстройку все возможности вашего приложения. Следует сосредоточиться на наиболее часто используемых и уместных действиях в контексте содержимого Outlook. Продумайте призыв к действию и убедитесь, что пользователь четко понимает, что ему делать, когда откроется область задач.</span><span class="sxs-lookup"><span data-stu-id="776a2-p104">An add-in should not attempt to do everything your app does. The focus should be on the most frequently used, and appropriate, actions in the context of Outlook content. Think about your call to action and make it clear what the user should do when your task pane opens.</span></span>

2. <span data-ttu-id="776a2-118">**Согласуйте надстройку с платформой**</span><span class="sxs-lookup"><span data-stu-id="776a2-118">**Make it feel as native as possible**</span></span>

   <span data-ttu-id="776a2-p105">В надстройке должны использоваться шаблоны, характерные для той платформы, на которой работает Outlook. Для этого следует соблюдать рекомендации по взаимодействию и внешнему виду для каждой платформы. Для приложения Outlook также есть свои рекомендации, которые важно учитывать. Хорошо спроектированная надстройка органично вписывается в интерфейс, платформу и Outlook.</span><span class="sxs-lookup"><span data-stu-id="776a2-p105">Your add-in should be designed using patterns native to the platform that Outlook is running on. To achieve this, be sure to respect and implement the interaction and visual guidelines set forth by each platform. Outlook has its own guidelines and those are also important to consider. A well-designed add-in will be an appropriate blend of your experience, the platform, and Outlook.</span></span>

   <span data-ttu-id="776a2-p106">Это означает, что ваша надстройка будет визуально отличаться при запуске в Outlook для iOS и в Outlook для Android. Рекомендуем ознакомиться с [Framework7](https://framework7.io/) как с одним из решений для оформления надстройки.</span><span class="sxs-lookup"><span data-stu-id="776a2-p106">This does mean that your add-in will have to visually be different when it runs in Outlook on iOS versus Android. We recommend taking a look at [Framework7](https://framework7.io/) as one option to help you with styling.</span></span>

3. <span data-ttu-id="776a2-125">**Сделайте работу приятной и проявите внимание к мелочам**</span><span class="sxs-lookup"><span data-stu-id="776a2-125">**Make it enjoyable to use and get the details right**</span></span>

   <span data-ttu-id="776a2-p107">Людям нравится пользоваться функциональными и красивыми продуктами. Вы можете гарантировать успех своей надстройке, уделив пристальное внимание каждому аспекту работы и визуального оформления. Действия, необходимые для выполнения задачи, должны быть понятными и логичными. В идеале каждое действие должно вызываться не более чем за два щелчка.</span><span class="sxs-lookup"><span data-stu-id="776a2-p107">People enjoy using products that are both functionally and visually appealing. You can help ensure the success of your add-in by crafting an experience where you've carefully considered every interaction and visual detail. The necessary steps to complete a task must be clear and relevant. Ideally, no action should be further than a click or two away.</span></span> 
   
   <span data-ttu-id="776a2-130">Старайтесь, чтобы пользователь не покидал контекст для выполнения действия.</span><span class="sxs-lookup"><span data-stu-id="776a2-130">Try not to take a user out of context to complete an action.</span></span> <span data-ttu-id="776a2-131">Пользователю должно быть легко заходить в надстройку и покидать ее, чтобы продолжить работу.</span><span class="sxs-lookup"><span data-stu-id="776a2-131">A user should easily be able to get in and out of your add-in and back to whatever she was doing before.</span></span> <span data-ttu-id="776a2-132">Надстройка не предназначена для того, чтобы проводить в ней много времени,&mdash;это лишь расширение базовых функций.</span><span class="sxs-lookup"><span data-stu-id="776a2-132">An add-in is not meant to be a destination to spend a lot of time in&mdash;it is an enhancement to our core functionality.</span></span> <span data-ttu-id="776a2-133">Правильно разработанная надстройка поможет повысить эффективность работы пользователей.</span><span class="sxs-lookup"><span data-stu-id="776a2-133">If done properly, your add-in will help us deliver on the goal of making people more productive.</span></span>

4. <span data-ttu-id="776a2-134">**Будьте осторожны с фирменной символикой**</span><span class="sxs-lookup"><span data-stu-id="776a2-134">**Brand wisely**</span></span>

   <span data-ttu-id="776a2-135">Мы ценим хороший фирменный стиль и понимаем, что важно обеспечить для пользователей уникальный процесс работы.</span><span class="sxs-lookup"><span data-stu-id="776a2-135">We value great branding, and we know it is important to provide users with your unique experience.</span></span> <span data-ttu-id="776a2-136">Но мы считаем, что лучший способ гарантировать успех своей надстройки — создать интуитивно понятный интерфейс, включающий ненавязчивые элементы фирменного стиля, а не нагружать его бросающейся в глаза фирменной символикой, которая только отвлекает пользователя и мешает навигации по системе.</span><span class="sxs-lookup"><span data-stu-id="776a2-136">But we feel the best way to ensure your add-in's success is to build an intuitive experience that subtly incorporates elements of your brand versus displaying persistent or obtrusive brand elements that only distract a user from moving through your system in an unencumbered manner.</span></span> 
    
   <span data-ttu-id="776a2-137">Чтобы удачно внедрить свой фирменный стиль, можно использовать фирменные цвета, значки и голос&mdash;при условии, что они не противоречат предпочитаемым шаблонам для платформы и требованиям к специальным возможностям.</span><span class="sxs-lookup"><span data-stu-id="776a2-137">A good way to incorporate your brand in a meaningful way is through the use of your brand colors, icons, and voice&mdash;assuming these don't conflict with the preferred platform patterns or accessibility requirements.</span></span> <span data-ttu-id="776a2-138">Стремитесь, чтобы основное внимание уделялось содержимому и выполнению задач, а не символике.</span><span class="sxs-lookup"><span data-stu-id="776a2-138">Strive to keep the focus on content and task completion, not brand attention.</span></span> 
    
   > [!NOTE]
   >  <span data-ttu-id="776a2-139">Объявления нельзя показывать в надстройках на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="776a2-139">Ads should not be shown within add-ins on iOS or Android.</span></span>

## <a name="design-patterns"></a><span data-ttu-id="776a2-140">Шаблоны проектирования</span><span class="sxs-lookup"><span data-stu-id="776a2-140">Design patterns</span></span>

> [!NOTE]
> <span data-ttu-id="776a2-141">Вышеперечисленные принципы применимы ко всем конечным точкам и платформам, но представленные ниже шаблоны и примеры относятся к мобильным надстройкам на платформе iOS.</span><span class="sxs-lookup"><span data-stu-id="776a2-141">While the above principles apply to all endpoints/platforms, the following patterns and examples are specific to mobile add-ins on the iOS platform.</span></span>

<span data-ttu-id="776a2-p111">Создать профессиональную надстройку вам помогут [шаблоны](../design/ux-design-pattern-templates.md), содержащие элементы мобильных приложений для iOS, которые работают в среде Outlook Mobile. Используя эти шаблоны, вы гарантируете, что надстройка органично впишется как в платформу iOS, так и в Outlook Mobile. Ниже представлены подробные описания этих шаблонов. Хотя этот список не является исчерпывающим, он положит начало библиотеке, которую мы продолжим развивать по мере обнаружения новых парадигм, которые партнеры пожелают добавлять в свои надстройки.</span><span class="sxs-lookup"><span data-stu-id="776a2-p111">To help you create a well-designed add-in, we have [templates](../design/ux-design-pattern-templates.md) that contain iOS mobile patterns that work within the Outlook Mobile environment. Leveraging these specific patterns will help ensure your add-in feels native to both the iOS platform and Outlook Mobile. These patterns are also detailed below. While not exhaustive, this is the start of a library that we will continue to build upon as we uncover additional paradigms partners wish to include in their add-ins.</span></span>  

### <a name="overview"></a><span data-ttu-id="776a2-146">Обзор</span><span class="sxs-lookup"><span data-stu-id="776a2-146">Overview</span></span>

<span data-ttu-id="776a2-147">Типичная надстройка состоит из следующих компонентов:</span><span class="sxs-lookup"><span data-stu-id="776a2-147">A typical add-in is made up of the following components.</span></span>

![Схема основных вариантов графического дизайна области задач на iOS](../images/outlook-mobile-design-overview.png)

![Схема основных вариантов графического дизайна области задач на Android](../images/outlook-mobile-design-overview-android.jpg)

### <a name="loading"></a><span data-ttu-id="776a2-150">Загрузка</span><span class="sxs-lookup"><span data-stu-id="776a2-150">Loading</span></span>

<span data-ttu-id="776a2-p112">Когда пользователь выбирает надстройку, ее пользовательский интерфейс должен отображаться как можно скорее. При наличии задержки используйте индикатор выполнения или индикатор работы. Индикатор выполнения следует использовать, если время загрузки можно определить, а индикатор работы — если оно непредсказуемо.</span><span class="sxs-lookup"><span data-stu-id="776a2-p112">When a user taps on your add-in, the UX should display as quickly as possible. If there is any delay, use a progress bar or activity indicator. A progress bar should be used when the amount of time is determinable and an activity indicator should be used when the amount of time is indeterminable.</span></span>

<span data-ttu-id="776a2-154">**Пример загрузки страниц на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-154">**An example of loading pages on iOS**</span></span>

![Примеры индикаторов выполнения и работы на iOS](../images/outlook-mobile-design-loading.png)

<span data-ttu-id="776a2-156">**Пример загрузки страниц на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-156">**An example of loading pages on Android**</span></span>

![Примеры индикаторов выполнения и работы на Android](../images/outlook-mobile-design-loading-android.jpg)


### <a name="sign-insign-up"></a><span data-ttu-id="776a2-158">Вход и регистрация</span><span class="sxs-lookup"><span data-stu-id="776a2-158">Sign in/Sign up</span></span>

<span data-ttu-id="776a2-159">Сделайте вход и регистрацию понятными и простыми.</span><span class="sxs-lookup"><span data-stu-id="776a2-159">Make your sign in (and sign up) flow straightforward and simple to use.</span></span>

<span data-ttu-id="776a2-160">**Пример страницы входа и регистрации на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-160">**An example sign in and sign up page on iOS**</span></span>

![Примеры страниц входа и регистрации на iOS](../images/outlook-mobile-design-signin.png)

<span data-ttu-id="776a2-162">**Пример страницы входа на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-162">**An example sign in page on Android**</span></span>

![Примеры страниц входа на Android](../images/outlook-mobile-design-signin-android.png)

### <a name="brand-bar"></a><span data-ttu-id="776a2-164">Панель с фирменной символикой</span><span class="sxs-lookup"><span data-stu-id="776a2-164">Brand bar</span></span>

<span data-ttu-id="776a2-p113">На первом экране надстройки должен отображаться элемент фирменной символики. Панель с фирменной символикой не только делает надстройку узнаваемой, но и создает контекст для пользователя. Так как панель навигации содержит название компании или торговой марки, их необязательно повторять на последующих страницах.</span><span class="sxs-lookup"><span data-stu-id="776a2-p113">The first screen of your add-in should include your branding element. Designed for recognition, the brand bar also helps set context for the user. Because the navigation bar contains the name of your company/brand, it's unnecessary to repeat the brand bar on subsequent pages.</span></span>

<span data-ttu-id="776a2-168">**Пример фирменной символики на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-168">**An example of branding on iOS**</span></span>

![Примеры панели с фирменной символикой на iOS](../images/outlook-mobile-design-branding.png)

<span data-ttu-id="776a2-170">**Пример фирменной символики на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-170">**An example of branding on Android**</span></span>

![Примеры панели с фирменной символикой на Android](../images/outlook-mobile-design-branding-android.png)

### <a name="margins"></a><span data-ttu-id="776a2-172">Поля</span><span class="sxs-lookup"><span data-stu-id="776a2-172">Margins</span></span>

<span data-ttu-id="776a2-173">Размер полей на мобильных устройствах должен составлять 15 пикселей (8 % экрана) с каждой стороны на iOS и 16 пикселей с каждой стороны на Android.</span><span class="sxs-lookup"><span data-stu-id="776a2-173">Mobile margins should be set to 15px (8% of screen) for each side, to align with Outlook iOS and 16px for each side to align with Outlook Android.</span></span>

![Примеры полей на iOS](../images/outlook-mobile-design-margins.png)

### <a name="typography"></a><span data-ttu-id="776a2-175">Шрифтовое оформление</span><span class="sxs-lookup"><span data-stu-id="776a2-175">Typography</span></span>

<span data-ttu-id="776a2-176">Оформление согласуется с оформлением Outlook на iOS и оптимизировано для беглого просмотра.</span><span class="sxs-lookup"><span data-stu-id="776a2-176">Typography usage is aligned to Outlook iOS and is kept simple for scannability.</span></span>

<span data-ttu-id="776a2-177">**Шрифтовое оформление на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-177">**Typography on iOS**</span></span>

![Примеры шрифтового оформления для iOS](../images/outlook-mobile-design-typography.png)

<span data-ttu-id="776a2-179">**Шрифтовое оформление на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-179">**Typography on Android**</span></span>

![Примеры шрифтового оформления для Android](../images/outlook-mobile-design-typography-android.png)

### <a name="color-palette"></a><span data-ttu-id="776a2-181">Цветовая палитра</span><span class="sxs-lookup"><span data-stu-id="776a2-181">Color palette</span></span>

<span data-ttu-id="776a2-p114">В Outlook iOS используется ненавязчивая цветовая схема.  Для согласованности рекомендуем использовать цвета только для действий и ошибок, а уникальные цвета использовать только на панели с фирменной символикой.</span><span class="sxs-lookup"><span data-stu-id="776a2-p114">Color usage is subtle in Outlook iOS.  To align, we ask that usage of color is localized to actions and error states, with only the brand bar using a unique color.</span></span>

![Цветовая палитра для iOS](../images/outlook-mobile-design-color-palette.png)

### <a name="cells"></a><span data-ttu-id="776a2-185">Ячейки</span><span class="sxs-lookup"><span data-stu-id="776a2-185">Cells</span></span>

<span data-ttu-id="776a2-186">Так как помечать страницы с помощью панели навигации невозможно, для этого следует использовать заголовки разделов.</span><span class="sxs-lookup"><span data-stu-id="776a2-186">Since the navigation bar cannot be used to label a page, use section titles to label pages.</span></span>

<span data-ttu-id="776a2-187">**Примеры ячеек на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-187">**Examples of cells on iOS**</span></span>

![Типы ячеек для iOS](../images/outlook-mobile-design-cell-types.png)
* * *
![Примеры правильного оформления ячеек для iOS](../images/outlook-mobile-design-cell-dos.png)
* * *
![Примеры неправильного оформления ячеек для iOS](../images/outlook-mobile-design-cell-donts.png)
* * *
![Ячейки и способы ввода для iOS](../images/outlook-mobile-design-cell-input.png)

<span data-ttu-id="776a2-192">**Примеры ячеек на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-192">**Examples of cells on Android**</span></span>

![Типы ячеек для Android](../images/outlook-mobile-design-cell-type-android.png)
* * *
![Примеры правильного оформления ячеек для Android](../images/outlook-mobile-design-cell-dos-android.png)
* * *
![Примеры неправильного оформления ячеек для Android](../images/outlook-mobile-design-cell-donts-android.png)
* * *
![Ячейки и способы ввода для Android, часть 1](../images/outlook-mobile-design-cell-input-1-android.png)

![Ячейки и способы ввода для Android, часть 2](../images/outlook-mobile-design-cell-input-2-android.png)

### <a name="actions"></a><span data-ttu-id="776a2-198">Действия</span><span class="sxs-lookup"><span data-stu-id="776a2-198">Actions</span></span>

<span data-ttu-id="776a2-199">Даже если ваше приложение выполняет множество задач, выберите самые важные действия для надстройки и сосредоточьтесь на них.</span><span class="sxs-lookup"><span data-stu-id="776a2-199">Even if your app handles a multitude of actions, think about the most important ones you want your add-in to perform, and concentrate on those.</span></span>

<span data-ttu-id="776a2-200">**Примеры действий на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-200">**Examples of actions on iOS**</span></span>

![Действия и ячейки в iOS](../images/outlook-mobile-design-action-cells.png)
* * *
![Примеры правильного выбора действий для iOS](../images/outlook-mobile-design-action-dos.png)

<span data-ttu-id="776a2-203">**Примеры действий на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-203">**Examples of actions on Android**</span></span>

![Действия и ячейки в Android](../images/outlook-mobile-design-action-cells-android.png)
* * *
![Примеры правильного выбора действий для Android](../images/outlook-mobile-design-action-dos-android.png)

### <a name="buttons"></a><span data-ttu-id="776a2-206">Кнопки</span><span class="sxs-lookup"><span data-stu-id="776a2-206">Buttons</span></span>

<span data-ttu-id="776a2-207">Кнопки используются в тех случаях, когда под ними есть другие элементы пользовательского интерфейса (в отличие от действий, которые должны быть последним элементом на экране).</span><span class="sxs-lookup"><span data-stu-id="776a2-207">Buttons are used when there are other UX elements below (vs. actions, where the action is the last element on the screen).</span></span>

<span data-ttu-id="776a2-208">**Примеры кнопок на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-208">**Examples of buttons on iOS**</span></span>

![Примеры кнопок для iOS](../images/outlook-mobile-design-buttons.png)

<span data-ttu-id="776a2-210">**Примеры кнопок на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-210">**Examples of buttons on Android**</span></span>

![Примеры кнопок для Android](../images/outlook-mobile-design-buttons-android.png)

### <a name="tabs"></a><span data-ttu-id="776a2-212">Вкладки</span><span class="sxs-lookup"><span data-stu-id="776a2-212">Tabs</span></span>

<span data-ttu-id="776a2-213">Вкладки помогают упорядочивать содержимое.</span><span class="sxs-lookup"><span data-stu-id="776a2-213">Tabs can aid in content organization.</span></span>

<span data-ttu-id="776a2-214">**Примеры вкладок на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-214">**Examples of tabs on iOS**</span></span>

![Примеры вкладок для iOS](../images/outlook-mobile-design-tabs.png)

<span data-ttu-id="776a2-216">**Примеры вкладок на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-216">**Examples of tabs on Android**</span></span>

![Примеры вкладок для Android](../images/outlook-mobile-design-tabs-android.png)

### <a name="icons"></a><span data-ttu-id="776a2-218">Значки</span><span class="sxs-lookup"><span data-stu-id="776a2-218">Icons</span></span>

<span data-ttu-id="776a2-p115">По мере возможности значки должны соответствовать текущему оформлению Outlook для iOS. Используйте стандартные размер и цвет.</span><span class="sxs-lookup"><span data-stu-id="776a2-p115">Icons should follow the current Outlook iOS design when possible. Use our standard size and color.</span></span>

<span data-ttu-id="776a2-221">**Примеры значков на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-221">**Examples of icons on iOS**</span></span>

![Примеры значков для iOS](../images/outlook-mobile-design-icons.png)

<span data-ttu-id="776a2-223">**Примеры значков на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-223">**Examples of icons on Android**</span></span>

![Примеры значков для Android](../images/outlook-mobile-design-icons-android.jpg)

## <a name="end-to-end-examples"></a><span data-ttu-id="776a2-225">Полные примеры</span><span class="sxs-lookup"><span data-stu-id="776a2-225">End-to-end examples</span></span>

<span data-ttu-id="776a2-226">Для выпуска первой версии надстроек Outlook Mobile мы тесно сотрудничали с нашими партнерами, занимающимися разработкой надстроек. Чтобы продемонстрировать потенциал надстроек для Outlook Mobile, наш дизайнер создал полноценные интерфейсы для каждой надстройки, используя наши рекомендации и шаблоны.</span><span class="sxs-lookup"><span data-stu-id="776a2-226">For our v1 Outlook Mobile Add-ins launch, we worked closely with our partners who were building add-ins. As a way to showcase the potential of their add-ins on Outlook Mobile, our designer put together end-to-end flows for each add-in, leveraging our guidelines and patterns.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="776a2-227">Эти примеры призваны показать идеальный подход к интерфейсу и визуальному оформлению надстройки и могут не полностью соответствовать функциям опубликованных версий надстроек.</span><span class="sxs-lookup"><span data-stu-id="776a2-227">These examples are meant to highlight the ideal way to approach both the interaction and visual design of an add-in and may not match the exact feature sets in the shipped versions of the add-ins.</span></span> 

### <a name="giphy"></a><span data-ttu-id="776a2-228">GIPHY</span><span class="sxs-lookup"><span data-stu-id="776a2-228">GIPHY</span></span>

<span data-ttu-id="776a2-229">**Пример GIPHY на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-229">**An example of GIPHY on iOS**</span></span>

![Полное оформление надстройки GIPHY на iOS](../images/outlook-mobile-design-giphy.png)

<span data-ttu-id="776a2-231">**Пример GIPHY на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-231">**An example of GIPHY on Android**</span></span>

![Полное оформление надстройки GIPHY на Android](../images/outlook-mobile-design-giphy-android.png)

### <a name="nimble"></a><span data-ttu-id="776a2-233">Nimble</span><span class="sxs-lookup"><span data-stu-id="776a2-233">Nimble</span></span>

<span data-ttu-id="776a2-234">**Пример Nimble на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-234">**An example of Nimble on iOS**</span></span>

![Полное оформление надстройки Nimble на iOS](../images/outlook-mobile-design-nimble.png)

<span data-ttu-id="776a2-236">**Пример Nimble на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-236">**An example of Nimble on Android**</span></span>

![Полное оформление надстройки Nimble на Android](../images/outlook-mobile-design-nimble-android.png)

### <a name="trello"></a><span data-ttu-id="776a2-238">Trello</span><span class="sxs-lookup"><span data-stu-id="776a2-238">Trello</span></span>

<span data-ttu-id="776a2-239">**Пример Trello на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-239">**An example of Trello on iOS**</span></span>

![Полное оформление надстройки Trello на iOS (часть 1)](../images/outlook-mobile-design-trello-1.png)
* * *
![Полное оформление надстройки Trello на iOS (часть 2)](../images/outlook-mobile-design-trello-2.png)
* * *
![Полное оформление надстройки Trello на iOS (часть 3)](../images/outlook-mobile-design-trello-3.png)

<span data-ttu-id="776a2-243">**Пример Trello на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-243">**An example of Trello on Android**</span></span>

![Полное оформление надстройки Trello на Android (часть 1)](../images/outlook-mobile-design-trello-1-android.png)
* * *
![Полное оформление надстройки Trello на Android (часть 2)](../images/outlook-mobile-design-trello-2-android.png)

### <a name="dynamics-crm"></a><span data-ttu-id="776a2-246">Dynamics CRM</span><span class="sxs-lookup"><span data-stu-id="776a2-246">Dynamics CRM</span></span>

<span data-ttu-id="776a2-247">**Пример Dynamics CRM на iOS**</span><span class="sxs-lookup"><span data-stu-id="776a2-247">**An example of Dynamics CRM on iOS**</span></span>

![Полное оформление надстройки Dynamics CRM на iOS](../images/outlook-mobile-design-crm.png)

<span data-ttu-id="776a2-249">**Пример Dynamics CRM на Android**</span><span class="sxs-lookup"><span data-stu-id="776a2-249">**An example of Dynamics CRM on Android**</span></span>

![Полное оформление надстройки Dynamics CRM на Android](../images/outlook-mobile-design-crm-android.png)
