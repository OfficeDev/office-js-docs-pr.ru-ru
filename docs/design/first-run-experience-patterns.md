---
title: Шаблоны интерфейса первого запуска для надстроек Office
description: Узнайте о лучших практиках разработки первого запуска в Office надстройки.
ms.date: 06/26/2018
localization_priority: Normal
ms.openlocfilehash: d020a281aca10805ba8fd1176403f3788f6d716c
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076345"
---
# <a name="first-run-experience-patterns"></a><span data-ttu-id="66046-103">Шаблоны интерфейса первого запуска</span><span class="sxs-lookup"><span data-stu-id="66046-103">First-run experience patterns</span></span>

<span data-ttu-id="66046-104">Интерфейс первого запуска (FRE) обеспечивает знакомство пользователя с вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="66046-104">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="66046-105">Когда пользователь впервые открывает надстройку, отображается интерфейс FRE, который дает им представление о функциях, возможностях и/или преимуществах надстройки.</span><span class="sxs-lookup"><span data-stu-id="66046-105">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="66046-106">Этот интерфейс формирует первое впечатление от надстройки и может сильно повлиять на вероятность того, что пользователь вернется и продолжит пользоваться вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="66046-106">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="66046-107">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="66046-107">Best practices</span></span>

<span data-ttu-id="66046-108">Следуйте этим рекомендациям при создании интерфейса первого запуска:</span><span class="sxs-lookup"><span data-stu-id="66046-108">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="66046-109">Правильно</span><span class="sxs-lookup"><span data-stu-id="66046-109">Do</span></span>|<span data-ttu-id="66046-110">Неправильно</span><span class="sxs-lookup"><span data-stu-id="66046-110">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="66046-111">Ясно и кратко опишите основные действия в надстройке.</span><span class="sxs-lookup"><span data-stu-id="66046-111">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="66046-112">Не указывайте сведения, не относящиеся к началу работы.</span><span class="sxs-lookup"><span data-stu-id="66046-112">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="66046-113">Предоставьте пользователям возможность выполнить действие, которое создаст у них положительное впечатление от использования надстройки.</span><span class="sxs-lookup"><span data-stu-id="66046-113">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="66046-114">Не следует ожидать, что пользователи изучат все возможности сразу.</span><span class="sxs-lookup"><span data-stu-id="66046-114">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="66046-115">Сосредоточьтесь на самом ценном действии.</span><span class="sxs-lookup"><span data-stu-id="66046-115">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="66046-116">Создайте привлекательный интерфейс, в котором пользователи захотят выполнить все действия.</span><span class="sxs-lookup"><span data-stu-id="66046-116">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="66046-117">Не заставляйте пользователей просматривать весь интерфейс первого запуска.</span><span class="sxs-lookup"><span data-stu-id="66046-117">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="66046-118">Предоставьте пользователям возможность обойти его.</span><span class="sxs-lookup"><span data-stu-id="66046-118">Give users an option to bypass the first-run experience.</span></span> |

<span data-ttu-id="66046-119">Решите, как часто необходимо применять интерфейс, используемый при первом запуске: один раз или периодически.</span><span class="sxs-lookup"><span data-stu-id="66046-119">Consider whether showing users the first-run experience once or periodically is important to your scenario.</span></span> <span data-ttu-id="66046-120">Например, если ваша надстройка используется только время от времени, пользователи могут забывать ее возможности, и тогда им будет полезно еще раз ознакомиться с интерфейсом первого запуска.</span><span class="sxs-lookup"><span data-stu-id="66046-120">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>

<span data-ttu-id="66046-121">При создании или улучшении интерфейса первого запуска для надстройки применяйте указанные ниже шаблоны.</span><span class="sxs-lookup"><span data-stu-id="66046-121">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>

## <a name="carousel"></a><span data-ttu-id="66046-122">Карусель</span><span class="sxs-lookup"><span data-stu-id="66046-122">Carousel</span></span>

<span data-ttu-id="66046-123">Карусель знакомит пользователей с рядом функций или предоставляет определенные сведения, прежде чем они начнут использовать надстройку.</span><span class="sxs-lookup"><span data-stu-id="66046-123">The carousel takes users through a series of features or informational pages before they start using the add-in.</span></span>

<span data-ttu-id="66046-124">*Рис. 1. Разрешить пользователям заранее или пропустить начало страниц потока карусель*</span><span class="sxs-lookup"><span data-stu-id="66046-124">*Figure 1. Allow users to advance or skip the beginning pages of the carousel flow*</span></span>

![Иллюстрация, показывающая шаг 1 карусели в первом запуске области задач Office настольного приложения.](../images/add-in-FRE-step-1.png)

<span data-ttu-id="66046-127">*Рис. 2. Свести к минимуму количество экранов карусель только до того, что необходимо для эффективного сообщения вашего сообщения*</span><span class="sxs-lookup"><span data-stu-id="66046-127">*Figure 2. Minimize the number of carousel screens to only what is needed to effectively communicate your message*</span></span>

![Иллюстрация, показывающая шаг 2 карусели в первом запуске области задач Office настольного приложения.](../images/add-in-FRE-step-2.png)

<span data-ttu-id="66046-130">*Рис. 3. Предоставление четкого вызова действий для выхода из первого запуска*</span><span class="sxs-lookup"><span data-stu-id="66046-130">*Figure 3. Provide a clear call to action to exit the first-run-experience*</span></span>

![Иллюстрация, показывающая шаг 3 карусели в первом запуске области задач Office настольного приложения.](../images/add-in-FRE-step-3.png)

## <a name="value-placemat"></a><span data-ttu-id="66046-133">Представление ценности</span><span class="sxs-lookup"><span data-stu-id="66046-133">Value Placemat</span></span>

<span data-ttu-id="66046-134">Представление ценности — это ценностное предложение вашей надстройки: размещение логотипа, ясно сформулированное ценностное предложение, краткое описание или обзор функций, а также призыв к действию.</span><span class="sxs-lookup"><span data-stu-id="66046-134">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>

<span data-ttu-id="66046-135">*Рис. 4. Placemat значения с логотипом, предложением четкого значения, сводка функций и вызов к действию*</span><span class="sxs-lookup"><span data-stu-id="66046-135">*Figure 4. A value placemat with logo, clear value proposition, feature summary, and call-to-action*</span></span>

![Иллюстрация, показывающая placemat значения в первом опытом запуска Office области задач настольных приложений.](../images/add-in-FRE-value.png)

### <a name="video-placemat"></a><span data-ttu-id="66046-138">Представление видео</span><span class="sxs-lookup"><span data-stu-id="66046-138">Video Placemat</span></span>

<span data-ttu-id="66046-139">Представление видео показывает пользователям видеоролик перед тем, как они начнут использовать вашу надстройку.</span><span class="sxs-lookup"><span data-stu-id="66046-139">The video placemat shows users a video before they start using your add-in.</span></span>

<span data-ttu-id="66046-140">*Рис. 5. Первый запуск видео-placemat — экран содержит изображение из видео с кнопкой воспроизведения и кнопкой "Вызов к действию"*</span><span class="sxs-lookup"><span data-stu-id="66046-140">*Figure 5. First run video placemat - The screen contains a still image from the video with a play button and clear call-to-action button*</span></span>

![Иллюстрация, показывающая видео-placemat в первом опытом запуска Office области задач настольного приложения.](../images/add-in-FRE-video.png)

<span data-ttu-id="66046-142">*Рис. 6. Video player — Пользователи, представленные с видео в диалоговом окне*</span><span class="sxs-lookup"><span data-stu-id="66046-142">*Figure 6. Video player - Users presented with a video within a dialog window*</span></span>

![Иллюстрация, показывающая видео в диалоговом окне с Office настольного приложения и области задач надстройки в фоновом режиме.](../images/add-in-FRE-video-dialog.png)
