---
title: Шаблоны интерфейса первого запуска для надстроек Office
description: Ознакомьтесь с рекомендациями по проектированию возможностей первого запуска в надстройках Office.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: c0528e869dd8ee7fe779785fb1a9b6d347deab75
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292955"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="e9d7f-103">Шаблоны интерфейса первого запуска</span><span class="sxs-lookup"><span data-stu-id="e9d7f-103">First-run experience patterns</span></span>

<span data-ttu-id="e9d7f-104">Интерфейс первого запуска (FRE) обеспечивает знакомство пользователя с вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-104">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="e9d7f-105">Когда пользователь впервые открывает надстройку, отображается интерфейс FRE, который дает им представление о функциях, возможностях и/или преимуществах надстройки.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-105">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="e9d7f-106">Этот интерфейс формирует первое впечатление от надстройки и может сильно повлиять на вероятность того, что пользователь вернется и продолжит пользоваться вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-106">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="e9d7f-107">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="e9d7f-107">Best practices</span></span>


<span data-ttu-id="e9d7f-108">Следуйте этим рекомендациям при создании интерфейса первого запуска:</span><span class="sxs-lookup"><span data-stu-id="e9d7f-108">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="e9d7f-109">Правильно</span><span class="sxs-lookup"><span data-stu-id="e9d7f-109">Do</span></span>|<span data-ttu-id="e9d7f-110">Неправильно</span><span class="sxs-lookup"><span data-stu-id="e9d7f-110">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="e9d7f-111">Ясно и кратко опишите основные действия в надстройке.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-111">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="e9d7f-112">Не указывайте сведения, не относящиеся к началу работы.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-112">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="e9d7f-113">Предоставьте пользователям возможность выполнить действие, которое создаст у них положительное впечатление от использования надстройки.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-113">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="e9d7f-114">Не следует ожидать, что пользователи изучат все возможности сразу.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-114">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="e9d7f-115">Сосредоточьтесь на самом ценном действии.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-115">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="e9d7f-116">Создайте привлекательный интерфейс, в котором пользователи захотят выполнить все действия.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-116">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="e9d7f-117">Не заставляйте пользователей просматривать весь интерфейс первого запуска.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-117">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="e9d7f-118">Предоставьте пользователям возможность обойти его.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-118">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="e9d7f-119">Решите, как часто необходимо применять интерфейс, используемый при первом запуске: один раз или периодически.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-119">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="e9d7f-120">Например, если ваша надстройка используется только время от времени, пользователи могут забывать ее возможности, и тогда им будет полезно еще раз ознакомиться с интерфейсом первого запуска.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-120">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="e9d7f-121">При создании или улучшении интерфейса первого запуска для надстройки применяйте указанные ниже шаблоны.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-121">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="e9d7f-122">Карусель</span><span class="sxs-lookup"><span data-stu-id="e9d7f-122">Carousel</span></span>


<span data-ttu-id="e9d7f-123">Карусель знакомит пользователей с рядом функций или предоставляет определенные сведения, прежде чем они начнут использовать надстройку.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-123">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="e9d7f-124">*Рисунок 1: Разрешите пользователям продвигать или пропускать начальные страницы для движения обойм.* 
 ![ Первый запуск — обойма 1 — спецификации для области задач рабочего стола](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="e9d7f-124">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel Step 1 - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="e9d7f-125">*На рисунке 2 показано, как уменьшить количество экранов обоймы, которые вы представляете пользователю, чтобы обеспечить эффективную передачу сообщения.* 
 ![ Первый запуск — обойма 2 — спецификации для области задач рабочего стола](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="e9d7f-125">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message.*
![First Run - Carousel Step 2 - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="e9d7f-126">*Рис. 3: вызов действия Clear для выхода из первого интерфейса при первом запуске.* 
 ![ Первый запуск — обойма 3 — спецификации для области задач рабочего стола](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="e9d7f-126">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel Step 3 - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="e9d7f-127">Представление ценности</span><span class="sxs-lookup"><span data-stu-id="e9d7f-127">Value Placemat</span></span>

<span data-ttu-id="e9d7f-128">Представление ценности — это ценностное предложение вашей надстройки: размещение логотипа, ясно сформулированное ценностное предложение, краткое описание или обзор функций, а также призыв к действию.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-128">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="e9d7f-129">![First Run – value представление — спецификации для области задач рабочего стола ](../images/add-in-FRE-value.png)
 *значение представление с логотипом, предложением Clear value, сводки функций и вызовом действия.*</span><span class="sxs-lookup"><span data-stu-id="e9d7f-129">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call-to-action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="e9d7f-130">Представление видео</span><span class="sxs-lookup"><span data-stu-id="e9d7f-130">Video Placemat</span></span>

<span data-ttu-id="e9d7f-131">Представление видео показывает пользователям видеоролик перед тем, как они начнут использовать вашу надстройку.</span><span class="sxs-lookup"><span data-stu-id="e9d7f-131">The video placemat shows users a video before they start using your add-in.</span></span>


<span data-ttu-id="e9d7f-132">*Рис. 1: первый запуск представление — на экране отображается изображение с изображением по-прежнему с помощью кнопки "воспроизводящий" и кнопка "очистить действие вызова".* 
 ![ Видео представление — спецификации для области задач рабочего стола](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="e9d7f-132">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call-to-action button.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="e9d7f-133">*Рис. 2: Video Player — в диалоговом окне отображаются видеоролики.* 
 ![ Видео представление — диалоговые окна — спецификации для области задач рабочего стола](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="e9d7f-133">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Dialog - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
