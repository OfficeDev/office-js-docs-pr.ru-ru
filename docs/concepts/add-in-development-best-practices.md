---
title: Рекомендации по разработке надстроек Office
description: Применяйте рекомендации при разработке для создания надстроек Office.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 17393d921129efcfb74eed3dd168633c2f58291b
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/18/2020
ms.locfileid: "49132181"
---
# <a name="best-practices-for-developing-office-add-ins"></a><span data-ttu-id="a965c-103">Рекомендации по разработке надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a965c-103">Best practices for developing Office Add-ins</span></span>

<span data-ttu-id="a965c-p101">Эффективные надстройки предоставляют уникальные и удобные функции, которые расширяют приложения Office, придавая им привлекательный внешний вид. Чтобы создать хорошую надстройку, сделайте работу пользователей удобной с первого запуска, разработайте первоклассный пользовательский интерфейс и оптимизируйте производительность надстройки. Применяя методики, описанные в этой статье, вы сможете создавать надстройки, которые помогают пользователям выполнять свои задачи быстро и эффективно.</span><span class="sxs-lookup"><span data-stu-id="a965c-p101">Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance. Apply the best practices described in this article to create add-ins that help your users complete their tasks quickly and efficiently.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="provide-clear-value"></a><span data-ttu-id="a965c-107">Преимущества должны быть очевидными</span><span class="sxs-lookup"><span data-stu-id="a965c-107">Provide clear value</span></span>

- <span data-ttu-id="a965c-p102">Создавайте надстройки, которые помогают пользователям выполнять свои задачи быстро и эффективно. Основное внимание следует уделить сценариям, применимым для приложений Office. Например:</span><span class="sxs-lookup"><span data-stu-id="a965c-p102">Create add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office applications. For example:</span></span>
  - <span data-ttu-id="a965c-111">обеспечьте более быстрое и простое выполнение основных задач разработки с меньшим количеством прерываний;</span><span class="sxs-lookup"><span data-stu-id="a965c-111">Make core authoring tasks faster and easier, with fewer interruptions.</span></span>
  - <span data-ttu-id="a965c-112">добавьте в Office новые сценарии;</span><span class="sxs-lookup"><span data-stu-id="a965c-112">Enable new scenarios within Office.</span></span>
  - <span data-ttu-id="a965c-113">Внедрять дополняющие службы в приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="a965c-113">Embed complementary services within Office applications.</span></span>
  - <span data-ttu-id="a965c-114">сделайте работу в Office более удобной, чтобы повысить производительность.</span><span class="sxs-lookup"><span data-stu-id="a965c-114">Improve the Office experience to enhance productivity.</span></span>
- <span data-ttu-id="a965c-115">Чтобы ваша надстройка быстро заинтересовала пользователей, обеспечьте [демонстрацию ее преимуществ уже при первом использовании](#create-an-engaging-first-run-experience).</span><span class="sxs-lookup"><span data-stu-id="a965c-115">Make sure that the value of your add-in is clear to users right away by [creating an engaging first run experience](#create-an-engaging-first-run-experience).</span></span>
- <span data-ttu-id="a965c-p103">Создайте [привлекательное описание надстройки в AppSource](/office/dev/store/create-effective-office-store-listings). Сделайте преимущества вашей надстройки очевидными из названия и описания. Не полагайтесь на то, что по вашей торговой марке будет понятно, для чего предназначена надстройка.</span><span class="sxs-lookup"><span data-stu-id="a965c-p103">Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings). Make the benefits of your add-in clear in your title and description. Don't rely on your brand to communicate what your add-in does.</span></span>

## <a name="create-an-engaging-first-run-experience"></a><span data-ttu-id="a965c-119">Удобство работы с первого запуска</span><span class="sxs-lookup"><span data-stu-id="a965c-119">Create an engaging first-run experience</span></span>

- <span data-ttu-id="a965c-p104">Привлекайте новых пользователей удобным и интуитивно понятным интерфейсом. Помните, что пользователи все еще решают, использовать вашу надстройку или забросить ее после загрузки из магазина.</span><span class="sxs-lookup"><span data-stu-id="a965c-p104">Engage new users with a highly usable and intuitive first experience. Note that users are still deciding whether to use or abandon an add-in after they download it from the store.</span></span>

- <span data-ttu-id="a965c-p105">Сделайте очевидными действия, необходимые для вызова вашей надстройки. Используйте видеоролики, указатели, панели разбиения на страницы и другие ресурсы, чтобы привлечь пользователей.</span><span class="sxs-lookup"><span data-stu-id="a965c-p105">Make the steps that the user needs to take to engage with your add-in clear. Use videos, placemats, paging panels, or other resources to entice users.</span></span>

- <span data-ttu-id="a965c-124">Если пользователям необходимо войти, чтобы использовать надстройку, следует донести до них ценность этой надстройки уже при ее запуске, а не просто просить их выполнить вход.</span><span class="sxs-lookup"><span data-stu-id="a965c-124">Reinforce the value proposition of your add-in on launch, rather than just asking users to sign in.</span></span>

- <span data-ttu-id="a965c-125">Разработайте обучающий интерфейс, чтобы помочь пользователям и персонализировать среду.</span><span class="sxs-lookup"><span data-stu-id="a965c-125">Provide teaching UI to guide users and make your UI personal.</span></span>

  ![Снимок экрана, иллюстрирующий сравнение "Do" и "не".](../images/contoso-part-catalog-do-dont.png)

- <span data-ttu-id="a965c-129">Если ваша контентная надстройка привязывается к данным в документе, включите пример данных или шаблон, чтобы показать пользователям рекомендуемый формат данных.</span><span class="sxs-lookup"><span data-stu-id="a965c-129">If your content add-in binds to data in the user's document, include sample data or a template to show users the data format to use.</span></span>

  ![Снимок экрана, иллюстрирующий сравнение "Do" и "не".](../images/add-in-title.png)

- <span data-ttu-id="a965c-p108">Предлагайте [бесплатные пробные версии](/office/dev/store/decide-on-a-pricing-model). Если для вашей надстройки требуется подписка, сделайте некоторые функции доступными без нее.</span><span class="sxs-lookup"><span data-stu-id="a965c-p108">Offer [free trials](/office/dev/store/decide-on-a-pricing-model). If your add-in requires a subscription, make some functionality available without a subscription.</span></span>

- <span data-ttu-id="a965c-p109">Упростите регистрацию. Используйте автоматическое заполнение сведений (адрес электронной почты, отображаемое имя) и пропустите проверку электронной почты.</span><span class="sxs-lookup"><span data-stu-id="a965c-p109">Make signup simple. Prefill information (email, display name) and skip email verifications.</span></span>

- <span data-ttu-id="a965c-p110">Избегайте всплывающих окон. Если вам необходимо их использовать, помогите пользователю включить всплывающее окно.</span><span class="sxs-lookup"><span data-stu-id="a965c-p110">Avoid pop ups. If you have to use them, guide the user to enable your pop up.</span></span>

<span data-ttu-id="a965c-139">Шаблоны, которые можно применять при разработке для первого запуска, представлены в статье [Конструктивные шаблоны для надстроек Office](../design/first-run-experience-patterns.md).</span><span class="sxs-lookup"><span data-stu-id="a965c-139">For patterns that you can apply as you develop your first-run experience, see [UX design patterns for Office Add-ins](../design/first-run-experience-patterns.md).</span></span>

## <a name="use-add-in-commands"></a><span data-ttu-id="a965c-140">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="a965c-140">Use add-in commands</span></span>

- <span data-ttu-id="a965c-p111">Предоставьте релевантные точки входа пользовательского интерфейса для надстройки, используя команды надстройки. Соответствующие сведения, включая рекомендации по оформлению, см. в статье [о командах надстроек](../design/add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="a965c-p111">Provide relevant UI entry points for your add-in by using add-in commands. For details, including design best practices, see [add-in commands](../design/add-in-commands.md).</span></span>

## <a name="apply-ux-design-principles"></a><span data-ttu-id="a965c-143">Принципы оформления элементов пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="a965c-143">Apply UX design principles</span></span>

- <span data-ttu-id="a965c-p112">Убедитесь, что внешний вид и поведение вашей надстройки согласованы с интерфейсом Office. Используйте [Office UI Fabric](https://developer.microsoft.com/fabric).</span><span class="sxs-lookup"><span data-stu-id="a965c-p112">Ensure that the look and feel and functionality of your add-in complements the Office experience. Use [Office UI Fabric](https://developer.microsoft.com/fabric).</span></span>

- <span data-ttu-id="a965c-p113">Больше содержимого, меньше хрома. Избегайте лишних элементов интерфейса, которые не представляют ценности для пользователя.</span><span class="sxs-lookup"><span data-stu-id="a965c-p113">Favor content over chrome. Avoid superfluous UI elements that don't add value to the user experience.</span></span>

- <span data-ttu-id="a965c-p114">Пользователь должен быть главным. Убедитесь, что пользователи понимают важные решения и могут с легкостью отменять действия, выполняемые надстройкой.</span><span class="sxs-lookup"><span data-stu-id="a965c-p114">Keep users in control. Ensure that users understand important decisions, and can easily reverse actions the add-in performs.</span></span>

- <span data-ttu-id="a965c-p115">Используйте фирменную символику как повод для доверия и ориентир. Она не должна слишком бросаться в глаза или служить рекламой.</span><span class="sxs-lookup"><span data-stu-id="a965c-p115">Use branding to inspire trust and orient users. Do not use branding to overwhelm or advertise to users.</span></span>

- <span data-ttu-id="a965c-p116">Избегайте прокрутки. Оптимизируйте надстройку для разрешения 1366 x 768.</span><span class="sxs-lookup"><span data-stu-id="a965c-p116">Avoid scrolling. Optimize for 1366 x 768 resolution.</span></span>

- <span data-ttu-id="a965c-154">Не включайте нелицензированные изображения.</span><span class="sxs-lookup"><span data-stu-id="a965c-154">Do not include unlicensed images.</span></span>

- <span data-ttu-id="a965c-155">Используйте [понятный и простой язык](../design/voice-guidelines.md) в надстройке.</span><span class="sxs-lookup"><span data-stu-id="a965c-155">Use [clear and simple language](../design/voice-guidelines.md) in your add-in.</span></span>

- <span data-ttu-id="a965c-156">Учитывайте специальные возможности. Обеспечьте удобство работы для всех пользователей и поддержку таких специальных возможностей, как средство чтения с экрана.</span><span class="sxs-lookup"><span data-stu-id="a965c-156">Account for accessibility - make your add-in easy for all users to interact with, and accommodate assistive technologies such as screen readers.</span></span>

- <span data-ttu-id="a965c-p117">Обеспечьте поддержку всех платформ и методов ввода, включая мышь, клавиатуру и [сенсорное управление](#optimize-for-touch). Убедитесь, что ваш пользовательский интерфейс поддерживает различные форм-факторы.</span><span class="sxs-lookup"><span data-stu-id="a965c-p117">Design for all platforms and input methods, including mouse/keyboard and [touch](#optimize-for-touch). Ensure that your UI is responsive to different form factors.</span></span>

### <a name="optimize-for-touch"></a><span data-ttu-id="a965c-159">Оптимизация для сенсорного управления</span><span class="sxs-lookup"><span data-stu-id="a965c-159">Optimize for touch</span></span>

- <span data-ttu-id="a965c-160">Используйте свойство [context. touchEnabled](/javascript/api/office/office.context#touchenabled) , чтобы определить, включено ли сенсорный ввод для приложения Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="a965c-160">Use the [Context.touchEnabled](/javascript/api/office/office.context#touchenabled) property to detect whether the Office application that your add-in runs on is touch enabled.</span></span>

  > [!NOTE]
  > <span data-ttu-id="a965c-161">Это свойство не поддерживается в Outlook.</span><span class="sxs-lookup"><span data-stu-id="a965c-161">This property is not supported in Outlook.</span></span>

- <span data-ttu-id="a965c-p118">Убедитесь, что размер всех элементов интерфейса удобен для сенсорного управления. Например, кнопки имеют достаточно большие размеры, а в полях ввода будет удобно вводить данные.</span><span class="sxs-lookup"><span data-stu-id="a965c-p118">Ensure that all controls are appropriately sized for touch interaction. For example, buttons have adequate touch targets, and input boxes are large enough for users to enter input.</span></span>

- <span data-ttu-id="a965c-164">Не рассчитывайте, что можно будет обойтись другими способами ввода, например наведением указателя или щелчком правой кнопкой мыши.</span><span class="sxs-lookup"><span data-stu-id="a965c-164">Do not rely on non-touch input methods like hover or right-click.</span></span>

- <span data-ttu-id="a965c-p119">Убедитесь, что надстройка работает как в книжной, так и в альбомной ориентации. Помните, что на сенсорных устройствах часть надстройки может быть закрыта экранной клавиатурой.</span><span class="sxs-lookup"><span data-stu-id="a965c-p119">Ensure that your add-in works in both portrait and landscape modes. Be aware that on touch devices, part of your add-in might be hidden by the soft keyboard.</span></span>

- <span data-ttu-id="a965c-167">Протестируйте надстройку на настоящем устройстве, выполнив [загрузку неопубликованного приложения ](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="a965c-167">Test your add-in on a real device by using [sideloading](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span></span>

> [!NOTE]
> <span data-ttu-id="a965c-168">Если используется [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric), то многие из элементов оформления настраиваются без вашего вмешательства.</span><span class="sxs-lookup"><span data-stu-id="a965c-168">If you're using [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) for your design elements, many of these elements are taken care of.</span></span>


## <a name="optimize-and-monitor-add-in-performance"></a><span data-ttu-id="a965c-169">Оптимизация и отслеживание производительности приложения</span><span class="sxs-lookup"><span data-stu-id="a965c-169">Optimize and monitor add-in performance</span></span>

- <span data-ttu-id="a965c-p120">Создайте ощущение быстрой реакции пользовательского интерфейса. Надстройка должна загружаться за 500 мс или меньше.</span><span class="sxs-lookup"><span data-stu-id="a965c-p120">Create the perception of fast UI responses. Your add-in should load in 500 ms or less.</span></span>

- <span data-ttu-id="a965c-172">Убедитесь, что все команды пользователя выполняются менее, чем за одну секунду.</span><span class="sxs-lookup"><span data-stu-id="a965c-172">Ensure that all user interactions respond in under one second.</span></span>

- <span data-ttu-id="a965c-173">Добавьте индикаторы загрузки для продолжительных операций.</span><span class="sxs-lookup"><span data-stu-id="a965c-173">Provide loading indicators for long-running operations.</span></span>

- <span data-ttu-id="a965c-p121">Используйте CDN для размещения изображений, ресурсов и общих библиотек. Загружайте как можно больше компонентов из одного источника.</span><span class="sxs-lookup"><span data-stu-id="a965c-p121">Use a CDN to host images, resources, and common libraries. Load as much as you can from one place.</span></span>

- <span data-ttu-id="a965c-p122">Соблюдайте стандартные методики для оптимизации веб-страницы. В производственной версии используйте только компактные версии библиотек. Загружайте только необходимые ресурсы и оптимизируйте их загрузку.</span><span class="sxs-lookup"><span data-stu-id="a965c-p122">Follow standard web practices to optimize your web page. In production, use only minified versions of libraries. Only load resources that you need, and optimize how resources are loaded.</span></span>

- <span data-ttu-id="a965c-p123">Если для выполнения операций требуется время, сообщайте об этом пользователям. Учитывайте пороговые значения, перечисленные в приведенной ниже таблице. Дополнительные сведения см. в статье [Ограничения ресурсов и оптимизация производительности надстроек Office](../concepts/resource-limits-and-performance-optimization.md).</span><span class="sxs-lookup"><span data-stu-id="a965c-p123">If operations take time to execute, provide feedback to users. Note the thresholds listed in the following table. For additional information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).</span></span>

  |<span data-ttu-id="a965c-182">Класс взаимодействия</span><span class="sxs-lookup"><span data-stu-id="a965c-182">Interaction class</span></span>|<span data-ttu-id="a965c-183">Target</span><span class="sxs-lookup"><span data-stu-id="a965c-183">Target</span></span>|<span data-ttu-id="a965c-184">Верхняя граница</span><span class="sxs-lookup"><span data-stu-id="a965c-184">Upper bound</span></span>|<span data-ttu-id="a965c-185">Человеческое восприятие</span><span class="sxs-lookup"><span data-stu-id="a965c-185">Human perception</span></span>|
  |:-----|:-----|:-----|:-----|
  |<span data-ttu-id="a965c-186">Мгновенно</span><span class="sxs-lookup"><span data-stu-id="a965c-186">Instant</span></span>|<span data-ttu-id="a965c-187"><=50 мс</span><span class="sxs-lookup"><span data-stu-id="a965c-187"><=50 ms</span></span>|<span data-ttu-id="a965c-188">100 мс</span><span class="sxs-lookup"><span data-stu-id="a965c-188">100 ms</span></span>|<span data-ttu-id="a965c-189">Без заметной задержки.</span><span class="sxs-lookup"><span data-stu-id="a965c-189">No noticeable delay.</span></span>|
  |<span data-ttu-id="a965c-190">Быстро</span><span class="sxs-lookup"><span data-stu-id="a965c-190">Fast</span></span>|<span data-ttu-id="a965c-191">50–100 мс</span><span class="sxs-lookup"><span data-stu-id="a965c-191">50-100 ms</span></span>|<span data-ttu-id="a965c-192">200 мс</span><span class="sxs-lookup"><span data-stu-id="a965c-192">200 ms</span></span>|<span data-ttu-id="a965c-p124">Минимально заметная задержка. Нет необходимости в информативном сопровождении.</span><span class="sxs-lookup"><span data-stu-id="a965c-p124">Minimally noticeable delay. No feedback necessary.</span></span>|
  |<span data-ttu-id="a965c-195">Нормальное</span><span class="sxs-lookup"><span data-stu-id="a965c-195">Typical</span></span>|<span data-ttu-id="a965c-196">100–300 мс</span><span class="sxs-lookup"><span data-stu-id="a965c-196">100-300 ms</span></span>|<span data-ttu-id="a965c-197">500 мс</span><span class="sxs-lookup"><span data-stu-id="a965c-197">500 ms</span></span>|<span data-ttu-id="a965c-p125">Достаточная скорость, но не более того. Нет необходимости в информативном сопровождении.</span><span class="sxs-lookup"><span data-stu-id="a965c-p125">Quick, but too slow to be described as fast. No feedback necessary.</span></span>|
  |<span data-ttu-id="a965c-200">Оперативно</span><span class="sxs-lookup"><span data-stu-id="a965c-200">Responsive</span></span>|<span data-ttu-id="a965c-201">300–500 мс</span><span class="sxs-lookup"><span data-stu-id="a965c-201">300-500 ms</span></span>|<span data-ttu-id="a965c-202">1 секунда</span><span class="sxs-lookup"><span data-stu-id="a965c-202">1 second</span></span>|<span data-ttu-id="a965c-p126">Не быстро, но надстройка реагирует хорошо. Нет необходимости в информативном сопровождении.</span><span class="sxs-lookup"><span data-stu-id="a965c-p126">Not fast, but still feels responsive. No feedback necessary.</span></span>|
  |<span data-ttu-id="a965c-205">Продолжительно</span><span class="sxs-lookup"><span data-stu-id="a965c-205">Continuous</span></span>|<span data-ttu-id="a965c-206">>500 мс</span><span class="sxs-lookup"><span data-stu-id="a965c-206">>500 ms</span></span>|<span data-ttu-id="a965c-207">5 секунд</span><span class="sxs-lookup"><span data-stu-id="a965c-207">5 seconds</span></span>|<span data-ttu-id="a965c-p127">Среднее время ожидания, надстройка реагирует не так хорошо. Может потребоваться информативное сопровождение.</span><span class="sxs-lookup"><span data-stu-id="a965c-p127">Medium wait, no longer feels responsive. Might need feedback.</span></span>|
  |<span data-ttu-id="a965c-210">Длительно</span><span class="sxs-lookup"><span data-stu-id="a965c-210">Captive</span></span>|<span data-ttu-id="a965c-211">>500 мс</span><span class="sxs-lookup"><span data-stu-id="a965c-211">>500 ms</span></span>|<span data-ttu-id="a965c-212">10 секунд</span><span class="sxs-lookup"><span data-stu-id="a965c-212">10 seconds</span></span>|<span data-ttu-id="a965c-p128">Длительная задержка, но не настолько, чтобы пользователь занялся чем-то другим. Может потребоваться информативное сопровождение.</span><span class="sxs-lookup"><span data-stu-id="a965c-p128">Long, but not long enough to do something else. Might need feedback.</span></span>|
  |<span data-ttu-id="a965c-215">Долго</span><span class="sxs-lookup"><span data-stu-id="a965c-215">Extended</span></span>|<span data-ttu-id="a965c-216">>500 мс</span><span class="sxs-lookup"><span data-stu-id="a965c-216">>500 ms</span></span>|<span data-ttu-id="a965c-217">Более 10 секунд</span><span class="sxs-lookup"><span data-stu-id="a965c-217">>10 seconds</span></span>|<span data-ttu-id="a965c-p129">Длительная задержка, при которой пользователь может заняться чем-то другим. Может потребоваться информативное сопровождение.</span><span class="sxs-lookup"><span data-stu-id="a965c-p129">Long enough to do something else while waiting. Might need feedback.</span></span>|
  |<span data-ttu-id="a965c-220">Слишком долго</span><span class="sxs-lookup"><span data-stu-id="a965c-220">Long running</span></span>|<span data-ttu-id="a965c-221">>5 с</span><span class="sxs-lookup"><span data-stu-id="a965c-221">>5 seconds</span></span>|<span data-ttu-id="a965c-222">>1 минуты</span><span class="sxs-lookup"><span data-stu-id="a965c-222">>1 minute</span></span>|<span data-ttu-id="a965c-223">Пользователи наверняка будут заниматься чем-то другим.</span><span class="sxs-lookup"><span data-stu-id="a965c-223">Users will certainly do something else.</span></span>|

- <span data-ttu-id="a965c-224">Отслеживайте работоспособность службы и используйте телеметрию для отслеживания успешной работы пользователя.</span><span class="sxs-lookup"><span data-stu-id="a965c-224">Monitor your service health, and use telemetry to monitor user success.</span></span>

- <span data-ttu-id="a965c-225">Минимизируйте обмен данными между надстройкой и документом Office.</span><span class="sxs-lookup"><span data-stu-id="a965c-225">Minimize data exchanges between the add-in and the Office document.</span></span> <span data-ttu-id="a965c-226">Дополнительные сведения см. [в статье Избегайте использования метода Context. Sync в циклах](correlated-objects-pattern.md).</span><span class="sxs-lookup"><span data-stu-id="a965c-226">For more information, see [Avoid using the context.sync method in loops](correlated-objects-pattern.md).</span></span>

## <a name="market-your-add-in"></a><span data-ttu-id="a965c-227">Маркетинг</span><span class="sxs-lookup"><span data-stu-id="a965c-227">Market your add-in</span></span>

- <span data-ttu-id="a965c-p131">Опубликуйте надстройку в [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) и [рекламируйте ее](/office/dev/store/promote-your-office-store-solution) на своем веб-сайте. Создайте [эффективное описание для AppSource](/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="a965c-p131">Publish your add-in to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) and [promote it](/office/dev/store/promote-your-office-store-solution) from your website. Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings).</span></span>

- <span data-ttu-id="a965c-p132">Давайте надстройкам лаконичные и информативные названия. Их длина не должна превышать 128 символов.</span><span class="sxs-lookup"><span data-stu-id="a965c-p132">Use succinct and descriptive add-in titles. Include no more than 128 characters.</span></span>

- <span data-ttu-id="a965c-p133">Составьте краткие и привлекательные описания надстройки. Дайте ответ на вопрос "Какую проблему решает эта надстройка?"</span><span class="sxs-lookup"><span data-stu-id="a965c-p133">Write short, compelling descriptions of your add-in. Answer the question "What problem does this add-in solve?".</span></span>

- <span data-ttu-id="a965c-p134">Опишите преимущества надстройки в названии и описании. Не полагайтесь на свою торговую марку.</span><span class="sxs-lookup"><span data-stu-id="a965c-p134">Convey the value proposition of your add-in in your title and description. Don't rely on your brand.</span></span>

- <span data-ttu-id="a965c-236">Создайте веб-сайт, который поможет пользователям найти и использовать вашу надстройку.</span><span class="sxs-lookup"><span data-stu-id="a965c-236">Create a website to help users find and use your add-in.</span></span>

## <a name="use-javascript-that-supports-internet-explorer"></a><span data-ttu-id="a965c-237">Использование JavaScript, поддерживающего Internet Explorer</span><span class="sxs-lookup"><span data-stu-id="a965c-237">Use JavaScript that supports Internet Explorer</span></span>

[!INCLUDE [How to support IE](../includes/es5-support.md)]

## <a name="see-also"></a><span data-ttu-id="a965c-238">См. также</span><span class="sxs-lookup"><span data-stu-id="a965c-238">See also</span></span>

- [<span data-ttu-id="a965c-239">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a965c-239">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="a965c-240">Сведения о программе для разработчиков Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="a965c-240">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
