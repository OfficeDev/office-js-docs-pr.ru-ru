# <a name="getstarted-element"></a><span data-ttu-id="167aa-101">Элемент GetStarted</span><span class="sxs-lookup"><span data-stu-id="167aa-101">GetStarted element</span></span>

<span data-ttu-id="167aa-p101">Предоставляет сведения для выноски, которая отображается при установке надстройки в ведущих приложениях Word, Excel, PowerPoint и OneNote. Элемент **GetStarted** является дочерним для элемента [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="167aa-p101">Provides information used by the callout that appears when the add-in is installed in Word, Excel, PowerPoint, and OneNote hosts. The **GetStarted** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span>

## <a name="child-elements"></a><span data-ttu-id="167aa-104">Дочерние элементы</span><span class="sxs-lookup"><span data-stu-id="167aa-104">Child elements</span></span>

| <span data-ttu-id="167aa-105">Элемент</span><span class="sxs-lookup"><span data-stu-id="167aa-105">Element</span></span>                       | <span data-ttu-id="167aa-106">Обязательный</span><span class="sxs-lookup"><span data-stu-id="167aa-106">Required</span></span> | <span data-ttu-id="167aa-107">Description</span><span class="sxs-lookup"><span data-stu-id="167aa-107">Description</span></span>                                        |
|:------------------------------|:--------:|:---------------------------------------------------|
| [<span data-ttu-id="167aa-108">Title</span><span class="sxs-lookup"><span data-stu-id="167aa-108">Title</span></span>](#title)               | <span data-ttu-id="167aa-109">Да</span><span class="sxs-lookup"><span data-stu-id="167aa-109">Yes</span></span>      | <span data-ttu-id="167aa-110">Определяет, где предоставляются функции надстройки.</span><span class="sxs-lookup"><span data-stu-id="167aa-110">Defines where an add-in exposes functionality.</span></span>     |
| [<span data-ttu-id="167aa-111">Description</span><span class="sxs-lookup"><span data-stu-id="167aa-111">Description</span></span>](#description)   | <span data-ttu-id="167aa-112">Да</span><span class="sxs-lookup"><span data-stu-id="167aa-112">Yes</span></span>      | <span data-ttu-id="167aa-113">URL-адрес файла, содержащего функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="167aa-113">A URL to a file that contains JavaScript functions.</span></span>|
| [<span data-ttu-id="167aa-114">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="167aa-114">LearnMoreUrl</span></span>](#learnmoreurl) | <span data-ttu-id="167aa-115">Нет</span><span class="sxs-lookup"><span data-stu-id="167aa-115">No</span></span>       | <span data-ttu-id="167aa-116">URL-адрес страницы с подробным описанием надстройки.</span><span class="sxs-lookup"><span data-stu-id="167aa-116">A URL to a page that explains the add-in in detail.</span></span>   |

### <a name="title"></a><span data-ttu-id="167aa-117">Title</span><span class="sxs-lookup"><span data-stu-id="167aa-117">Title</span></span> 

<span data-ttu-id="167aa-p102">Обязательный. Заголовок в верхней части выноски. Атрибут **resid** ссылается на допустимый идентификатор элемента **ShortStrings** в разделе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="167aa-p102">Required. The title used for the top of the callout. The **resid** attribute references a valid ID in the **ShortStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="description"></a><span data-ttu-id="167aa-121">Description</span><span class="sxs-lookup"><span data-stu-id="167aa-121">Description</span></span>

<span data-ttu-id="167aa-p103">Обязательный.  Атрибут **resid** ссылается на допустимый идентификатор элемента **ShortStrings** в разделе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="167aa-p103">Required. The description / body content for the callout. The **resid** attribute references a valid ID in the **LongStrings** element in the [Resources](resources.md) section.</span></span>

### <a name="learnmoreurl"></a><span data-ttu-id="167aa-125">LearnMoreUrl</span><span class="sxs-lookup"><span data-stu-id="167aa-125">LearnMoreUrl</span></span>

<span data-ttu-id="167aa-p104">Обязательный. URL-адрес страницы, где пользователь может узнать больше о надстройке. Атрибут **resid** ссылается на допустимый идентификатор элемента **Urls** в разделе [Resources](resources.md).</span><span class="sxs-lookup"><span data-stu-id="167aa-p104">Required. The URL to a page where the user can learn more about your add-in. The **resid** attribute references a valid ID in the **Urls** element in the [Resources](resources.md) section.</span></span>

> [!NOTE]
> <span data-ttu-id="167aa-129">В настоящее время элемент **LearnMoreUrl** не отображается в клиентах Word, Excel и PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="167aa-129">NOTE:**LearnMoreUrl** does not currently render in Word, Excel, or PowerPoint clients.</span></span> <span data-ttu-id="167aa-130">Рекомендуем добавить URL-адрес всех клиентов, чтобы этот адрес отображался, когда он станет доступен.</span><span class="sxs-lookup"><span data-stu-id="167aa-130">We recommend that you add this URL for all clients so that the URL will render when it becomes available.</span></span> 

## <a name="see-also"></a><span data-ttu-id="167aa-131">См. также</span><span class="sxs-lookup"><span data-stu-id="167aa-131">See also</span></span>

<span data-ttu-id="167aa-132">В следующих примерах кода используется элемент **GetStarted**:</span><span class="sxs-lookup"><span data-stu-id="167aa-132">The following code samples use the **GetStarted** element:</span></span>

* [<span data-ttu-id="167aa-133">Веб-надстройка Excel для работы с форматированием таблиц и диаграмм</span><span class="sxs-lookup"><span data-stu-id="167aa-133">Excel Web Add-in for Manipulating Table and Chart Formatting</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker)
* [<span data-ttu-id="167aa-134">JavaScript SpecKit для надстроек Word</span><span class="sxs-lookup"><span data-stu-id="167aa-134">Word Add-in JavaScript SpecKit</span></span>](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
* [<span data-ttu-id="167aa-135">Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint</span><span class="sxs-lookup"><span data-stu-id="167aa-135">Insert Excel charts using Microsoft Graph in a PowerPoint Add-in</span></span>](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
