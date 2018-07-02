# <a name="first-run-experience-patterns"></a><span data-ttu-id="cf29c-101">Шаблоны экрана первого запуска</span><span class="sxs-lookup"><span data-stu-id="cf29c-101">First-run experience patterns</span></span>

<span data-ttu-id="cf29c-102">Экран первого запуска (FRE) обеспечивает знакомство пользователя с вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="cf29c-102">A First-run Experience (FRE) is a user's introduction to your add-in.</span></span> <span data-ttu-id="cf29c-103">Когда пользователь впервые открывает надстройку, отображается экран FRE, который дает им представление о функциях, возможностях и/или преимуществах надстройки.</span><span class="sxs-lookup"><span data-stu-id="cf29c-103">An FRE is presented when a user opens an add-in for the first time and provides them with insight into the functions, features, and/or benefits of the add-in.</span></span> <span data-ttu-id="cf29c-104">Этот экран формирует первое впечатление от надстройки и может сильно повлиять на вероятность того, что пользователь вернется и продолжит пользоваться вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="cf29c-104">This experience helps shape the user's impression of an add-in and can strongly influence their likelihood to come back to and continue using your add-in..</span></span>

## <a name="best-practices"></a><span data-ttu-id="cf29c-105">Советы и рекомендации</span><span class="sxs-lookup"><span data-stu-id="cf29c-105">Best practices</span></span>


<span data-ttu-id="cf29c-106">Следуйте этим рекомендациям при создании экрана первого запуска:</span><span class="sxs-lookup"><span data-stu-id="cf29c-106">Follow these best practices when crafting your first-run experience:</span></span>

|<span data-ttu-id="cf29c-107">Правильно</span><span class="sxs-lookup"><span data-stu-id="cf29c-107">Do</span></span>|<span data-ttu-id="cf29c-108">Неправильно</span><span class="sxs-lookup"><span data-stu-id="cf29c-108">Don't</span></span>|
|:------|:------|
|<span data-ttu-id="cf29c-109">Ясно и кратко опишите основные действия в надстройке.</span><span class="sxs-lookup"><span data-stu-id="cf29c-109">Provide a simple and brief introduction to the main actions in the add-in.</span></span> | <span data-ttu-id="cf29c-110">Включите в описание только те данные и элементы интерфейса, которые имеют отношение к началу работы.</span><span class="sxs-lookup"><span data-stu-id="cf29c-110">Don't include information and call-outs that are not relevant to getting started.</span></span>
|<span data-ttu-id="cf29c-111">Предоставьте пользователям возможность выполнить действие, которое создаст у них положительное впечатление от использования надстройки.</span><span class="sxs-lookup"><span data-stu-id="cf29c-111">Give users the opportunity to complete an action that will positively impact their use of the add-in.</span></span> | <span data-ttu-id="cf29c-112">Не следует ожидать, что пользователи изучат все возможности сразу.</span><span class="sxs-lookup"><span data-stu-id="cf29c-112">Don't expect users to learn everything at once.</span></span> <span data-ttu-id="cf29c-113">Сосредоточьтесь на самом ценном действии.</span><span class="sxs-lookup"><span data-stu-id="cf29c-113">Focus on the action that provides the most value.</span></span>
|<span data-ttu-id="cf29c-114">Создайте интересное описание, которое пользователи захотят дочитать до конца.</span><span class="sxs-lookup"><span data-stu-id="cf29c-114">Create an engaging experience that users will want to complete.</span></span> | <span data-ttu-id="cf29c-115">Не заставляйте пользователей просматривать весь экран первого запуска.</span><span class="sxs-lookup"><span data-stu-id="cf29c-115">Don't force the users to click through the first-run experience.</span></span> <span data-ttu-id="cf29c-116">Предоставьте пользователям возможность обойти его.</span><span class="sxs-lookup"><span data-stu-id="cf29c-116">Give users an option to bypass the first-run experience.</span></span> |



<span data-ttu-id="cf29c-117">Решите, как часто для вашего сценария необходимо отображать экран первого запуска: однократно или периодически.</span><span class="sxs-lookup"><span data-stu-id="cf29c-117">Consider whether showing users the first-run experience once or many times is important to your scenario.</span></span> <span data-ttu-id="cf29c-118">Например, если ваша надстройка используется только время от времени, пользователи могут забывать ее возможности, и тогда им будет полезно еще раз ознакомиться с экраном первого запуска.</span><span class="sxs-lookup"><span data-stu-id="cf29c-118">For example, if your add-in is only utilized periodically, users may become less familiar with your add-in and may benefit from another interaction with the first-run experience.</span></span>



<span data-ttu-id="cf29c-119">При создании или улучшении экрана первого запуска для вашей надстройки применяйте следующие шаблоны.</span><span class="sxs-lookup"><span data-stu-id="cf29c-119">Apply the following patterns as applicable to create or enhance the first-run experience for your add-in.</span></span>



## <a name="carousel"></a><span data-ttu-id="cf29c-120">Карусель</span><span class="sxs-lookup"><span data-stu-id="cf29c-120">Carousel</span></span>


<span data-ttu-id="cf29c-121">Карусель знакомит пользователей с последовательностью функций или информационных страниц, прежде чем они начнут пользоваться надстройкой.</span><span class="sxs-lookup"><span data-stu-id="cf29c-121">Walkthrough takes users through a series of features or information before they start using the add-in. (PDF, code)</span></span>

<span data-ttu-id="cf29c-122">*Рисунок 1. Предоставьте пользователям возможность прокручивать или пропускать начальные страницы карусели.*
![Первый запуск — Карусель — Спецификации для области задач рабочего стола](../images/add-in-FRE-step-1.png)</span><span class="sxs-lookup"><span data-stu-id="cf29c-122">*Figure 1: Allow users to advance or skip the beginning pages of the carousel flow.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-1.png)</span></span>



<span data-ttu-id="cf29c-123">*Рисунок 2. Показывайте как можно меньше экранов карусели, оставьте только те, которые необходимы для эффективного взаимодействия с пользователем*
![Первый запуск — Карусель — Спецификации для области задач рабочего стола](../images/add-in-FRE-step-2.png)</span><span class="sxs-lookup"><span data-stu-id="cf29c-123">*Figure 2: Minimize the number of carousel screens you present to the user to only what is needed to effectively communicate your message*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-2.png)</span></span>


<span data-ttu-id="cf29c-124">*Рисунок 3. Предоставьте ясные указания по выходу из экрана первого запуска.*
![Первый запуск — Карусель — Спецификации для области задач рабочего стола](../images/add-in-FRE-step-3.png)</span><span class="sxs-lookup"><span data-stu-id="cf29c-124">*Figure 3: Provide a clear call to action to exit the first-run-experience.*
![First Run - Carousel - Specifications for desktop task pane](../images/add-in-FRE-step-3.png)</span></span>



## <a name="value-placemat"></a><span data-ttu-id="cf29c-125">Представление ценности</span><span class="sxs-lookup"><span data-stu-id="cf29c-125">Value Placemat</span></span>

<span data-ttu-id="cf29c-126">Представление ценности — это ценностное предложение вашей надстройки: размещение логотипа, ясно сформулированное ценностное предложение, краткое описание или сводка основных возможностей, а также призыв к действию.</span><span class="sxs-lookup"><span data-stu-id="cf29c-126">The value placement communicates your add-in's value proposition through logo placement, a clearly stated value proposition, feature highlights or summary, and a call-to-action.</span></span>



<span data-ttu-id="cf29c-127">![Первый запуск — Представление ценности — Спецификации для панели задач рабочего стола](../images/add-in-FRE-value.png)
*Представление ценности с логотипом, ясное ценностное предложение, сводка функций и призыв к действию.*</span><span class="sxs-lookup"><span data-stu-id="cf29c-127">![First Run - Value Placemat - Specifications for desktop task pane](../images/add-in-FRE-value.png)
*A value placemat with logo, clear value proposition, feature summary, and call to action.*</span></span>


### <a name="video-placemat"></a><span data-ttu-id="cf29c-128">Представление видео</span><span class="sxs-lookup"><span data-stu-id="cf29c-128">Video Placemat</span></span>

<span data-ttu-id="cf29c-129">Представление видео показывает пользователям видеоролик перед началом работы с вашей надстройкой.</span><span class="sxs-lookup"><span data-stu-id="cf29c-129">Video shows users a video before they start using your add-in. (spec, code)</span></span>


<span data-ttu-id="cf29c-130">*Рисунок 1. Представление первого запуска — Экран содержит кадр из видео с кнопкой запуска и кнопку с четким призывом к действию.*![Ознакомительное видео — Спецификации для области задач рабочего стола](../images/add-in-FRE-video.png)</span><span class="sxs-lookup"><span data-stu-id="cf29c-130">*Figure 1: First Run Placemat - The screen contains a still image from the video with a play button and clear call to action button.*![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video.png)</span></span>



<span data-ttu-id="cf29c-131">*Рисунок 2. Видеопроигрыватель — Пользователям представляется видео в диалоговом окне.*
![Представление видео — Спецификации для области задач рабочего стола](../images/add-in-FRE-video-dialog.png)</span><span class="sxs-lookup"><span data-stu-id="cf29c-131">*Figure 2: Video Player - Users are presented with a video within a dialog window.*
![Video Placemat - Specifications for desktop task pane](../images/add-in-FRE-video-dialog.png)</span></span>
