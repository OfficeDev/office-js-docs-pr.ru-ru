---
title: Рекомендации по разработке надстроек Office
description: Применяйте рекомендации при разработке для создания надстроек Office.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 71a090721768b822a058091c258dd0c6c87e1425
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890541"
---
# <a name="best-practices-for-developing-office-add-ins"></a><span data-ttu-id="63516-103">Рекомендации по разработке надстроек Office</span><span class="sxs-lookup"><span data-stu-id="63516-103">Best practices for developing Office Add-ins</span></span>

<span data-ttu-id="63516-p101">Эффективные надстройки предоставляют уникальные и удобные функции, которые расширяют приложения Office, придавая им привлекательный внешний вид. Чтобы создать хорошую надстройку, сделайте работу пользователей удобной с первого запуска, разработайте первоклассный пользовательский интерфейс и оптимизируйте производительность надстройки. Применяя методики, описанные в этой статье, вы сможете создавать надстройки, которые помогают пользователям выполнять свои задачи быстро и эффективно.</span><span class="sxs-lookup"><span data-stu-id="63516-p101">Effective add-ins offer unique and compelling functionality that extends Office applications in a visually appealing way. To create a great add-in, provide an engaging first-time experience for your users, design a first-class UI experience, and optimize your add-in's performance. Apply the best practices described in this article to create add-ins that help your users complete their tasks quickly and efficiently.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="provide-clear-value"></a><span data-ttu-id="63516-107">Преимущества должны быть очевидными</span><span class="sxs-lookup"><span data-stu-id="63516-107">Provide clear value</span></span>

- <span data-ttu-id="63516-p102">Создавайте надстройки, которые помогают пользователям выполнять свои задачи быстро и эффективно. Основное внимание следует уделить сценариям, применимым для приложений Office. Например:</span><span class="sxs-lookup"><span data-stu-id="63516-p102">Create add-ins that help users complete tasks quickly and efficiently. Focus on scenarios that make sense for Office applications. For example:</span></span>
 - <span data-ttu-id="63516-111">обеспечьте более быстрое и простое выполнение основных задач разработки с меньшим количеством прерываний;</span><span class="sxs-lookup"><span data-stu-id="63516-111">Make core authoring tasks faster and easier, with fewer interruptions.</span></span>
 - <span data-ttu-id="63516-112">добавьте в Office новые сценарии;</span><span class="sxs-lookup"><span data-stu-id="63516-112">Enable new scenarios within Office.</span></span>
 - <span data-ttu-id="63516-113">внедрите в ведущие приложения Office дополнительные службы;</span><span class="sxs-lookup"><span data-stu-id="63516-113">Embed complementary services within Office hosts.</span></span>
 - <span data-ttu-id="63516-114">сделайте работу в Office более удобной, чтобы повысить производительность.</span><span class="sxs-lookup"><span data-stu-id="63516-114">Improve the Office experience to enhance productivity.</span></span>
- <span data-ttu-id="63516-115">Чтобы ваша надстройка быстро заинтересовала пользователей, обеспечьте [демонстрацию ее преимуществ уже при первом использовании](#create-an-engaging-first-run-experience).</span><span class="sxs-lookup"><span data-stu-id="63516-115">Make sure that the value of your add-in is clear to users right away by [creating an engaging first run experience](#create-an-engaging-first-run-experience).</span></span>
- <span data-ttu-id="63516-p103">Создайте [привлекательное описание надстройки в AppSource](/office/dev/store/create-effective-office-store-listings). Сделайте преимущества вашей надстройки очевидными из названия и описания. Не полагайтесь на то, что по вашей торговой марке будет понятно, для чего предназначена надстройка.</span><span class="sxs-lookup"><span data-stu-id="63516-p103">Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings). Make the benefits of your add-in clear in your title and description. Don't rely on your brand to communicate what your add-in does.</span></span>


## <a name="create-an-engaging-first-run-experience"></a><span data-ttu-id="63516-119">Удобство работы с первого запуска</span><span class="sxs-lookup"><span data-stu-id="63516-119">Create an engaging first-run experience</span></span>

- <span data-ttu-id="63516-p104">Привлекайте новых пользователей удобным и интуитивно понятным интерфейсом. Помните, что пользователи все еще решают, использовать вашу надстройку или забросить ее после загрузки из магазина.</span><span class="sxs-lookup"><span data-stu-id="63516-p104">Engage new users with a highly usable and intuitive first experience. Note that users are still deciding whether to use or abandon an add-in after they download it from the store.</span></span>

- <span data-ttu-id="63516-p105">Сделайте очевидными действия, необходимые для вызова вашей надстройки. Используйте видеоролики, указатели, панели разбиения на страницы и другие ресурсы, чтобы привлечь пользователей.</span><span class="sxs-lookup"><span data-stu-id="63516-p105">Make the steps that the user needs to take to engage with your add-in clear. Use videos, placemats, paging panels, or other resources to entice users.</span></span>

- <span data-ttu-id="63516-124">Если пользователям необходимо войти, чтобы использовать надстройку, следует донести до них ценность этой надстройки уже при ее запуске, а не просто просить их выполнить вход.</span><span class="sxs-lookup"><span data-stu-id="63516-124">Reinforce the value proposition of your add-in on launch, rather than just asking users to sign in.</span></span>

- <span data-ttu-id="63516-125">Разработайте обучающий интерфейс, чтобы помочь пользователям и персонализировать среду.</span><span class="sxs-lookup"><span data-stu-id="63516-125">Provide teaching UI to guide users and make your UI personal.</span></span>

   ![Снимок экрана: область задач надстройки с начальными этапами работы рядом с надстройкой без этих этапов](../images/contoso-part-catalog-do-dont.png)

- <span data-ttu-id="63516-127">Если ваша контентная надстройка привязывается к данным в документе, включите пример данных или шаблон, чтобы показать пользователям рекомендуемый формат данных.</span><span class="sxs-lookup"><span data-stu-id="63516-127">If your content add-in binds to data in the user's document, include sample data or a template to show users the data format to use.</span></span>

   ![Снимок экрана: контентная надстройка с данными рядом с контентной надстройкой без данных](../images/add-in-title.png)

- <span data-ttu-id="63516-p106">Предлагайте [бесплатные пробные версии](/office/dev/store/decide-on-a-pricing-model). Если для вашей надстройки требуется подписка, сделайте некоторые функции доступными без нее.</span><span class="sxs-lookup"><span data-stu-id="63516-p106">Offer [free trials](/office/dev/store/decide-on-a-pricing-model). If your add-in requires a subscription, make some functionality available without a subscription.</span></span>

- <span data-ttu-id="63516-p107">Упростите регистрацию. Используйте автоматическое заполнение сведений (адрес электронной почты, отображаемое имя) и пропустите проверку электронной почты.</span><span class="sxs-lookup"><span data-stu-id="63516-p107">Make signup simple. Prefill information (email, display name) and skip email verifications.</span></span>

- <span data-ttu-id="63516-p108">Избегайте всплывающих окон. Если вам необходимо их использовать, помогите пользователю включить всплывающее окно.</span><span class="sxs-lookup"><span data-stu-id="63516-p108">Avoid pop ups. If you have to use them, guide the user to enable your pop up.</span></span>

<span data-ttu-id="63516-135">Шаблоны, которые можно применять при разработке для первого запуска, представлены в статье [Конструктивные шаблоны для надстроек Office](../design/first-run-experience-patterns.md).</span><span class="sxs-lookup"><span data-stu-id="63516-135">For patterns that you can apply as you develop your first-run experience, see [UX design patterns for Office Add-ins](../design/first-run-experience-patterns.md).</span></span>

## <a name="use-add-in-commands"></a><span data-ttu-id="63516-136">Команды надстроек</span><span class="sxs-lookup"><span data-stu-id="63516-136">Use add-in commands</span></span>

- <span data-ttu-id="63516-p109">Предоставьте релевантные точки входа пользовательского интерфейса для надстройки, используя команды надстройки. Соответствующие сведения, включая рекомендации по оформлению, см. в статье [о командах надстроек](../design/add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="63516-p109">Provide relevant UI entry points for your add-in by using add-in commands. For details, including design best practices, see [add-in commands](../design/add-in-commands.md).</span></span>

## <a name="apply-ux-design-principles"></a><span data-ttu-id="63516-139">Принципы оформления элементов пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="63516-139">Apply UX design principles</span></span>

- <span data-ttu-id="63516-p110">Убедитесь, что внешний вид и поведение вашей надстройки согласованы с интерфейсом Office. Используйте [Office UI Fabric](https://developer.microsoft.com/fabric).</span><span class="sxs-lookup"><span data-stu-id="63516-p110">Ensure that the look and feel and functionality of your add-in complements the Office experience. Use [Office UI Fabric](https://developer.microsoft.com/fabric).</span></span>

- <span data-ttu-id="63516-p111">Больше содержимого, меньше хрома. Избегайте лишних элементов интерфейса, которые не представляют ценности для пользователя.</span><span class="sxs-lookup"><span data-stu-id="63516-p111">Favor content over chrome. Avoid superfluous UI elements that don't add value to the user experience.</span></span>

- <span data-ttu-id="63516-p112">Пользователь должен быть главным. Убедитесь, что пользователи понимают важные решения и могут с легкостью отменять действия, выполняемые надстройкой.</span><span class="sxs-lookup"><span data-stu-id="63516-p112">Keep users in control. Ensure that users understand important decisions, and can easily reverse actions the add-in performs.</span></span>

- <span data-ttu-id="63516-p113">Используйте фирменную символику как повод для доверия и ориентир. Она не должна слишком бросаться в глаза или служить рекламой.</span><span class="sxs-lookup"><span data-stu-id="63516-p113">Use branding to inspire trust and orient users. Do not use branding to overwhelm or advertise to users.</span></span>

- <span data-ttu-id="63516-p114">Избегайте прокрутки. Оптимизируйте надстройку для разрешения 1366 x 768.</span><span class="sxs-lookup"><span data-stu-id="63516-p114">Avoid scrolling. Optimize for 1366 x 768 resolution.</span></span>

- <span data-ttu-id="63516-150">Не включайте нелицензированные изображения.</span><span class="sxs-lookup"><span data-stu-id="63516-150">Do not include unlicensed images.</span></span>

- <span data-ttu-id="63516-151">Используйте [понятный и простой язык](../design/voice-guidelines.md) в надстройке.</span><span class="sxs-lookup"><span data-stu-id="63516-151">Use [clear and simple language](../design/voice-guidelines.md) in your add-in.</span></span>

- <span data-ttu-id="63516-152">Учитывайте специальные возможности. Обеспечьте удобство работы для всех пользователей и поддержку таких специальных возможностей, как средство чтения с экрана.</span><span class="sxs-lookup"><span data-stu-id="63516-152">Account for accessibility - make your add-in easy for all users to interact with, and accommodate assistive technologies such as screen readers.</span></span>

- <span data-ttu-id="63516-p115">Обеспечьте поддержку всех платформ и методов ввода, включая мышь, клавиатуру и [сенсорное управление](#optimize-for-touch). Убедитесь, что ваш пользовательский интерфейс поддерживает различные форм-факторы.</span><span class="sxs-lookup"><span data-stu-id="63516-p115">Design for all platforms and input methods, including mouse/keyboard and [touch](#optimize-for-touch). Ensure that your UI is responsive to different form factors.</span></span>

### <a name="optimize-for-touch"></a><span data-ttu-id="63516-155">Оптимизация для сенсорного управления</span><span class="sxs-lookup"><span data-stu-id="63516-155">Optimize for touch</span></span>

- <span data-ttu-id="63516-156">Используйте свойство [Context.touchEnabled](/javascript/api/office/office.context), чтобы определить, поддерживается ли сенсорное управление ведущим приложением, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="63516-156">Use the [Context.touchEnabled](/javascript/api/office/office.context) property to detect whether the host application your add-in runs on is touch enabled.</span></span>

  > [!NOTE]
  > <span data-ttu-id="63516-157">Это свойство не поддерживается в Outlook.</span><span class="sxs-lookup"><span data-stu-id="63516-157">This property is not supported in Outlook.</span></span>

- <span data-ttu-id="63516-p116">Убедитесь, что размер всех элементов интерфейса удобен для сенсорного управления. Например, кнопки имеют достаточно большие размеры, а в полях ввода будет удобно вводить данные.</span><span class="sxs-lookup"><span data-stu-id="63516-p116">Ensure that all controls are appropriately sized for touch interaction. For example, buttons have adequate touch targets, and input boxes are large enough for users to enter input.</span></span>

- <span data-ttu-id="63516-160">Не рассчитывайте, что можно будет обойтись другими способами ввода, например наведением указателя или щелчком правой кнопкой мыши.</span><span class="sxs-lookup"><span data-stu-id="63516-160">Do not rely on non-touch input methods like hover or right-click.</span></span>

- <span data-ttu-id="63516-p117">Убедитесь, что надстройка работает как в книжной, так и в альбомной ориентации. Помните, что на сенсорных устройствах часть надстройки может быть закрыта экранной клавиатурой.</span><span class="sxs-lookup"><span data-stu-id="63516-p117">Ensure that your add-in works in both portrait and landscape modes. Be aware that on touch devices, part of your add-in might be hidden by the soft keyboard.</span></span>

- <span data-ttu-id="63516-163">Протестируйте надстройку на настоящем устройстве, выполнив [загрузку неопубликованного приложения ](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span><span class="sxs-lookup"><span data-stu-id="63516-163">Test your add-in on a real device by using [sideloading](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).</span></span>

> [!NOTE]
> <span data-ttu-id="63516-164">Если используется [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric), то многие из элементов оформления настраиваются без вашего вмешательства.</span><span class="sxs-lookup"><span data-stu-id="63516-164">If you're using [Office UI Fabric](https://github.com/OfficeDev/Office-UI-Fabric) for your design elements, many of these elements are taken care of.</span></span>


## <a name="optimize-and-monitor-add-in-performance"></a><span data-ttu-id="63516-165">Оптимизация и отслеживание производительности приложения</span><span class="sxs-lookup"><span data-stu-id="63516-165">Optimize and monitor add-in performance</span></span>

- <span data-ttu-id="63516-p118">Создайте ощущение быстрой реакции пользовательского интерфейса. Надстройка должна загружаться за 500 мс или меньше.</span><span class="sxs-lookup"><span data-stu-id="63516-p118">Create the perception of fast UI responses. Your add-in should load in 500 ms or less.</span></span>

- <span data-ttu-id="63516-168">Убедитесь, что все команды пользователя выполняются менее, чем за одну секунду.</span><span class="sxs-lookup"><span data-stu-id="63516-168">Ensure that all user interactions respond in under one second.</span></span>

-  <span data-ttu-id="63516-169">Добавьте индикаторы загрузки для продолжительных операций.</span><span class="sxs-lookup"><span data-stu-id="63516-169">Provide loading indicators for long-running operations.</span></span>

- <span data-ttu-id="63516-p119">Используйте CDN для размещения изображений, ресурсов и общих библиотек. Загружайте как можно больше компонентов из одного источника.</span><span class="sxs-lookup"><span data-stu-id="63516-p119">Use a CDN to host images, resources, and common libraries. Load as much as you can from one place.</span></span>

- <span data-ttu-id="63516-p120">Соблюдайте стандартные методики для оптимизации веб-страницы. В производственной версии используйте только компактные версии библиотек. Загружайте только необходимые ресурсы и оптимизируйте их загрузку.</span><span class="sxs-lookup"><span data-stu-id="63516-p120">Follow standard web practices to optimize your web page. In production, use only minified versions of libraries. Only load resources that you need, and optimize how resources are loaded.</span></span>

- <span data-ttu-id="63516-p121">Если для выполнения операций требуется время, сообщайте об этом пользователям. Учитывайте пороговые значения, перечисленные в приведенной ниже таблице. Дополнительные сведения см. в статье [Ограничения ресурсов и оптимизация производительности надстроек Office](../concepts/resource-limits-and-performance-optimization.md).</span><span class="sxs-lookup"><span data-stu-id="63516-p121">If operations take time to execute, provide feedback to users. Note the thresholds listed in the following table. For additional information, see [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md).</span></span>

  |<span data-ttu-id="63516-178">**Класс взаимодействия**</span><span class="sxs-lookup"><span data-stu-id="63516-178">**Interaction class**</span></span>|<span data-ttu-id="63516-179">**Целевой объект**</span><span class="sxs-lookup"><span data-stu-id="63516-179">**Target**</span></span>|<span data-ttu-id="63516-180">**Верхняя граница**</span><span class="sxs-lookup"><span data-stu-id="63516-180">**Upper bound**</span></span>|<span data-ttu-id="63516-181">**Впечатление от использования**</span><span class="sxs-lookup"><span data-stu-id="63516-181">**Human perception**</span></span>|
  |:-----|:-----|:-----|:-----|
  |<span data-ttu-id="63516-182">Мгновенно</span><span class="sxs-lookup"><span data-stu-id="63516-182">Instant</span></span>|<span data-ttu-id="63516-183"><=50 мс</span><span class="sxs-lookup"><span data-stu-id="63516-183"><=50 ms</span></span>|<span data-ttu-id="63516-184">100 мс</span><span class="sxs-lookup"><span data-stu-id="63516-184">100 ms</span></span>|<span data-ttu-id="63516-185">Без заметной задержки.</span><span class="sxs-lookup"><span data-stu-id="63516-185">No noticeable delay.</span></span>|
  |<span data-ttu-id="63516-186">Быстро</span><span class="sxs-lookup"><span data-stu-id="63516-186">Fast</span></span>|<span data-ttu-id="63516-187">50–100 мс</span><span class="sxs-lookup"><span data-stu-id="63516-187">50-100 ms</span></span>|<span data-ttu-id="63516-188">200 мс</span><span class="sxs-lookup"><span data-stu-id="63516-188">200 ms</span></span>|<span data-ttu-id="63516-p122">Минимально заметная задержка. Нет необходимости в информативном сопровождении.</span><span class="sxs-lookup"><span data-stu-id="63516-p122">Minimally noticeable delay. No feedback necessary.</span></span>|
  |<span data-ttu-id="63516-191">Нормальное</span><span class="sxs-lookup"><span data-stu-id="63516-191">Typical</span></span>|<span data-ttu-id="63516-192">100–300 мс</span><span class="sxs-lookup"><span data-stu-id="63516-192">100-300 ms</span></span>|<span data-ttu-id="63516-193">500 мс</span><span class="sxs-lookup"><span data-stu-id="63516-193">500 ms</span></span>|<span data-ttu-id="63516-p123">Достаточная скорость, но не более того. Нет необходимости в информативном сопровождении.</span><span class="sxs-lookup"><span data-stu-id="63516-p123">Quick, but too slow to be described as fast. No feedback necessary.</span></span>|
  |<span data-ttu-id="63516-196">Оперативно</span><span class="sxs-lookup"><span data-stu-id="63516-196">Responsive</span></span>|<span data-ttu-id="63516-197">300–500 мс</span><span class="sxs-lookup"><span data-stu-id="63516-197">300-500 ms</span></span>|<span data-ttu-id="63516-198">1 секунда</span><span class="sxs-lookup"><span data-stu-id="63516-198">1 second</span></span>|<span data-ttu-id="63516-p124">Не быстро, но надстройка реагирует хорошо. Нет необходимости в информативном сопровождении.</span><span class="sxs-lookup"><span data-stu-id="63516-p124">Not fast, but still feels responsive. No feedback necessary.</span></span>|
  |<span data-ttu-id="63516-201">Продолжительно</span><span class="sxs-lookup"><span data-stu-id="63516-201">Continuous</span></span>|<span data-ttu-id="63516-202">>500 мс</span><span class="sxs-lookup"><span data-stu-id="63516-202">>500 ms</span></span>|<span data-ttu-id="63516-203">5 секунд</span><span class="sxs-lookup"><span data-stu-id="63516-203">5 seconds</span></span>|<span data-ttu-id="63516-p125">Среднее время ожидания, надстройка реагирует не так хорошо. Может потребоваться информативное сопровождение.</span><span class="sxs-lookup"><span data-stu-id="63516-p125">Medium wait, no longer feels responsive. Might need feedback.</span></span>|
  |<span data-ttu-id="63516-206">Длительно</span><span class="sxs-lookup"><span data-stu-id="63516-206">Captive</span></span>|<span data-ttu-id="63516-207">>500 мс</span><span class="sxs-lookup"><span data-stu-id="63516-207">>500 ms</span></span>|<span data-ttu-id="63516-208">10 секунд</span><span class="sxs-lookup"><span data-stu-id="63516-208">10 seconds</span></span>|<span data-ttu-id="63516-p126">Длительная задержка, но не настолько, чтобы пользователь занялся чем-то другим. Может потребоваться информативное сопровождение.</span><span class="sxs-lookup"><span data-stu-id="63516-p126">Long, but not long enough to do something else. Might need feedback.</span></span>|
  |<span data-ttu-id="63516-211">Долго</span><span class="sxs-lookup"><span data-stu-id="63516-211">Extended</span></span>|<span data-ttu-id="63516-212">>500 мс</span><span class="sxs-lookup"><span data-stu-id="63516-212">>500 ms</span></span>|<span data-ttu-id="63516-213">Более 10 секунд</span><span class="sxs-lookup"><span data-stu-id="63516-213">>10 seconds</span></span>|<span data-ttu-id="63516-p127">Длительная задержка, при которой пользователь может заняться чем-то другим. Может потребоваться информативное сопровождение.</span><span class="sxs-lookup"><span data-stu-id="63516-p127">Long enough to do something else while waiting. Might need feedback.</span></span>|
  |<span data-ttu-id="63516-216">Слишком долго</span><span class="sxs-lookup"><span data-stu-id="63516-216">Long running</span></span>|<span data-ttu-id="63516-217">>5 с</span><span class="sxs-lookup"><span data-stu-id="63516-217">>5 seconds</span></span>|<span data-ttu-id="63516-218">>1 минуты</span><span class="sxs-lookup"><span data-stu-id="63516-218">>1 minute</span></span>|<span data-ttu-id="63516-219">Пользователи наверняка будут заниматься чем-то другим.</span><span class="sxs-lookup"><span data-stu-id="63516-219">Users will certainly do something else.</span></span>|

- <span data-ttu-id="63516-220">Отслеживайте работоспособность службы и используйте телеметрию для отслеживания успешной работы пользователя.</span><span class="sxs-lookup"><span data-stu-id="63516-220">Monitor your service health, and use telemetry to monitor user success.</span></span>


## <a name="market-your-add-in"></a><span data-ttu-id="63516-221">Маркетинг</span><span class="sxs-lookup"><span data-stu-id="63516-221">Market your add-in</span></span>

- <span data-ttu-id="63516-p128">Опубликуйте надстройку в [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) и [рекламируйте ее](/office/dev/store/promote-your-office-store-solution) на своем веб-сайте. Создайте [эффективное описание для AppSource](/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="63516-p128">Publish your add-in to [AppSource](/office/dev/store/submit-to-appsource-via-partner-center) and [promote it](/office/dev/store/promote-your-office-store-solution) from your website. Create an [effective AppSource listing](/office/dev/store/create-effective-office-store-listings).</span></span>

- <span data-ttu-id="63516-p129">Давайте надстройкам лаконичные и информативные названия. Их длина не должна превышать 128 символов.</span><span class="sxs-lookup"><span data-stu-id="63516-p129">Use succinct and descriptive add-in titles. Include no more than 128 characters.</span></span>

- <span data-ttu-id="63516-p130">Составьте краткие и привлекательные описания надстройки. Дайте ответ на вопрос "Какую проблему решает эта надстройка?"</span><span class="sxs-lookup"><span data-stu-id="63516-p130">Write short, compelling descriptions of your add-in. Answer the question "What problem does this add-in solve?".</span></span>

- <span data-ttu-id="63516-p131">Опишите преимущества надстройки в названии и описании. Не полагайтесь на свою торговую марку.</span><span class="sxs-lookup"><span data-stu-id="63516-p131">Convey the value proposition of your add-in in your title and description. Don't rely on your brand.</span></span>

- <span data-ttu-id="63516-230">Создайте веб-сайт, который поможет пользователям найти и использовать вашу надстройку.</span><span class="sxs-lookup"><span data-stu-id="63516-230">Create a website to help users find and use your add-in.</span></span>

## <a name="see-also"></a><span data-ttu-id="63516-231">См. также</span><span class="sxs-lookup"><span data-stu-id="63516-231">See also</span></span>

- [<span data-ttu-id="63516-232">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="63516-232">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
