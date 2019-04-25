---
title: Шаблоны интерфейса первого запуска для надстроек Office
description: ''
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: 85f8e4f7e0082e00ad5064333470f589e449af45
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32446856"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="2f02d-102">Шаблоны интерфейса первого запуска</span><span class="sxs-lookup"><span data-stu-id="2f02d-102">First-run experience patterns</span></span>

<span data-ttu-id="2f02d-103">Интерфейс первого запуска (FRE) обеспечивает знакомство пользователя с вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="2f02d-103">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="2f02d-104">Когда пользователь впервые открывает надстройку, отображается интерфейс FRE, который дает им представление о функциях, возможностях и/или преимуществах надстройки.</span><span class="sxs-lookup"><span data-stu-id="2f02d-104">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="2f02d-105">Этот интерфейс формирует первое впечатление от надстройки и может сильно повлиять на вероятность того, что пользователь вернется и продолжит пользоваться вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="2f02d-105">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="2f02d-106">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="2f02d-106">Best practices</span></span>


<span data-ttu-id="2f02d-107">Следуйте этим рекомендациям при создании интерфейса первого запуска:</span><span class="sxs-lookup"><span data-stu-id="2f02d-107">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="2f02d-108">Правильно</span><span class="sxs-lookup"><span data-stu-id="2f02d-108">Do</span></span>|<span data-ttu-id="2f02d-109">Неправильно</span><span class="sxs-lookup"><span data-stu-id="2f02d-109">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="2f02d-110">Ясно и кратко опишите основные действия в надстройке.</span><span class="sxs-lookup"><span data-stu-id="2f02d-110">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="2f02d-111">Не указывайте сведения, не относящиеся к началу работы.</span><span class="sxs-lookup"><span data-stu-id="2f02d-111">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="2f02d-112">Предоставьте пользователям возможность выполнить действие, которое создаст у них положительное впечатление от использования надстройки.</span><span class="sxs-lookup"><span data-stu-id="2f02d-112">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="2f02d-113">Не следует ожидать, что пользователи изучат все возможности сразу.</span><span class="sxs-lookup"><span data-stu-id="2f02d-113">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="2f02d-114">Сосредоточьтесь на самом ценном действии.</span><span class="sxs-lookup"><span data-stu-id="2f02d-114">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="2f02d-115">Создайте привлекательный интерфейс, в котором пользователи захотят выполнить все действия.</span><span class="sxs-lookup"><span data-stu-id="2f02d-115">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="2f02d-116">Не заставляйте пользователей просматривать весь интерфейс первого запуска.</span><span class="sxs-lookup"><span data-stu-id="2f02d-116">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="2f02d-117">Предоставьте пользователям возможность обойти его.</span><span class="sxs-lookup"><span data-stu-id="2f02d-117">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="2f02d-118">Решите, как часто необходимо применять интерфейс, используемый при первом запуске: один раз или периодически.</span><span class="sxs-lookup"><span data-stu-id="2f02d-118">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="2f02d-119">Например, если ваша надстройка используется только время от времени, пользователи могут забывать ее возможности, и тогда им будет полезно еще раз ознакомиться с интерфейсом первого запуска.</span><span class="sxs-lookup"><span data-stu-id="2f02d-119">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="2f02d-120">При создании или улучшении интерфейса первого запуска для надстройки применяйте указанные ниже шаблоны.</span><span class="sxs-lookup"><span data-stu-id="2f02d-120">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="2f02d-121">Карусель</span><span class="sxs-lookup"><span data-stu-id="2f02d-121">Carousel</span></span>


<span data-ttu-id="2f02d-122">Карусель знакомит пользователей с рядом функций или предоставляет определенные сведения, прежде чем они начнут использовать надстройку.</span><span class="sxs-lookup"><span data-stu-id="2f02d-122">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="2f02d-123">*Рисунок 1. Предоставьте пользователям возможность прокручивать или пропускать начальные страницы карусели.*
![Первый запуск — карусель: спецификации для области задач рабочего стола](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="2f02d-123">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="2f02d-124">*Рисунок 2. Показывайте как можно меньше экранов карусели, оставьте только те, которые необходимы для эффективного взаимодействия с пользователем*
![Первый запуск — карусель: спецификации для области задач рабочего стола](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="2f02d-124">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="2f02d-125">*Рисунок 3. Предоставьте ясные указания по выходу из интерфейса первого запуска.*
![Первый запуск — карусель: спецификации для области задач рабочего стола](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="2f02d-125">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="2f02d-126">Представление ценности</span><span class="sxs-lookup"><span data-stu-id="2f02d-126">Value Placemat</span></span>

<span data-ttu-id="2f02d-127">Представление ценности — это ценностное предложение вашей надстройки: размещение логотипа, ясно сформулированное ценностное предложение, краткое описание или обзор функций, а также призыв к действию.</span><span class="sxs-lookup"><span data-stu-id="2f02d-127">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="2f02d-128">![Первый запуск — представление ценности: спецификации для области задач рабочего стола](../images/add-in-FRE-value.png)
*Представление ценности с логотипом, ясное ценностное предложение, обзор функций и призыв к действию.*</span><span class="sxs-lookup"><span data-stu-id="2f02d-128">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call to action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="2f02d-129">Представление видео</span><span class="sxs-lookup"><span data-stu-id="2f02d-129">Video Placemat</span></span>

<span data-ttu-id="2f02d-130">Представление видео показывает пользователям видеоролик перед тем, как они начнут использовать вашу надстройку.</span><span class="sxs-lookup"><span data-stu-id="2f02d-130">The video placemat shows users a video before they start using your add-in.</span></span>


<span data-ttu-id="2f02d-131">*Рисунок 1. Представление первого запуска — экран содержит кадр из видео с кнопкой воспроизведения и кнопку с четким призывом к действию.*![Представление видео: спецификации для области задач рабочего стола](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="2f02d-131">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call to action button.*![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="2f02d-132">*Рисунок 2. Видеопроигрыватель — пользователям представляется видео в диалоговом окне.*
![Представление видео: спецификации для области задач рабочего стола](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="2f02d-132">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
