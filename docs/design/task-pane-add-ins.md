---
title: Области задач в надстройках Office
description: Области задач предоставляют пользователям доступ к элементам управления интерфейсом, которые выполняют код для изменения документов или сообщений электронной почты, а также для отображения данных из источника данных.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: d235d6c437ee124441389e68b54fc6ab8cde8dae
ms.sourcegitcommit: 30f6c620380075e3459cac748ca0c656427b384d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/12/2021
ms.locfileid: "52330152"
---
# <a name="task-panes-in-office-add-ins"></a><span data-ttu-id="1205e-103">Области задач в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="1205e-103">Task panes in Office Add-ins</span></span>

<span data-ttu-id="1205e-p101">Области задач — это области интерфейса, которые обычно отображаются в правой части окна Word, PowerPoint, Excel и Outlook. Элементы области задач выполняют код для изменения документов или писем, а также для отображения данных. Используйте области задач, когда вам не нужно внедрять функции прямо в документ.</span><span class="sxs-lookup"><span data-stu-id="1205e-p101">Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.</span></span>

<span data-ttu-id="1205e-107">*Рис. 1. Типичный макет области задач*</span><span class="sxs-lookup"><span data-stu-id="1205e-107">*Figure 1. Typical task pane layout*</span></span>

![Иллюстрация, отобразив типичную макет области задач с вкладками раздела в верхней части, логотипом компании и именем компании в левом нижнем ряду и значком параметров в правом нижнем справа](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a><span data-ttu-id="1205e-109">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="1205e-109">Best practices</span></span>

|<span data-ttu-id="1205e-110">Правильно</span><span class="sxs-lookup"><span data-stu-id="1205e-110">Do</span></span>|<span data-ttu-id="1205e-111">Неправильно</span><span class="sxs-lookup"><span data-stu-id="1205e-111">Don't</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="1205e-112">Включите имя надстройки в название.</span><span class="sxs-lookup"><span data-stu-id="1205e-112">Include the name of your add-in in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="1205e-113">Не включайте в него название вашей компании.</span><span class="sxs-lookup"><span data-stu-id="1205e-113">Don't append your company name to the title.</span></span></li></ul>|
|<ul><li><span data-ttu-id="1205e-114">Используйте короткие описательные имена в названии.</span><span class="sxs-lookup"><span data-stu-id="1205e-114">Use short descriptive names in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="1205e-115">Не добавляйте строки, такие как "надстройка", "для Word" или "для Office" к названию надстройки.</span><span class="sxs-lookup"><span data-stu-id="1205e-115">Don't append strings such as "add-in," "for Word," or "for Office" to the title of your add-in.</span></span></li></ul>|
|<ul><li><span data-ttu-id="1205e-116">Добавьте элемент навигации или управления, такой как CommandBar или Pivot, в верхнюю часть надстройки.</span><span class="sxs-lookup"><span data-stu-id="1205e-116">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span></li></ul>||
|<ul><li><span data-ttu-id="1205e-117">Включите элемент фирменной символики, такой как BrandBar, в нижнюю часть надстройки, если только она не будет использоваться исключительно в Outlook.</span><span class="sxs-lookup"><span data-stu-id="1205e-117">Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.</span></span></li></ul>||

## <a name="variants"></a><span data-ttu-id="1205e-118">Варианты</span><span class="sxs-lookup"><span data-stu-id="1205e-118">Variants</span></span>

<span data-ttu-id="1205e-119">На следующих изображениях покажут различные размеры области задач с лентой Приложение Office с разрешением 1366x768.</span><span class="sxs-lookup"><span data-stu-id="1205e-119">The following images show the various task pane sizes with the Office app ribbon at a 1366x768 resolution.</span></span> <span data-ttu-id="1205e-120">Чтобы вставить строку формул в Excel, требуется дополнительное пространство по вертикали.</span><span class="sxs-lookup"><span data-stu-id="1205e-120">For Excel, additional vertical space is required to accommodate the formula bar.</span></span>  

<span data-ttu-id="1205e-121">*Рис. 2. Размеры области задач в классических приложениях Office 2016*</span><span class="sxs-lookup"><span data-stu-id="1205e-121">*Figure 2. Office 2016 desktop task pane sizes*</span></span>

![Схема с отображением размеров области задач рабочего стола с разрешением 1366x768](../images/office-2016-taskpane-sizes.png)

- <span data-ttu-id="1205e-123">Excel — 320x455 пикселей</span><span class="sxs-lookup"><span data-stu-id="1205e-123">Excel - 320x455 pixels</span></span>
- <span data-ttu-id="1205e-124">PowerPoint — 320x531 пикселей</span><span class="sxs-lookup"><span data-stu-id="1205e-124">PowerPoint - 320x531 pixels</span></span>
- <span data-ttu-id="1205e-125">Word — 320x531 пикселей</span><span class="sxs-lookup"><span data-stu-id="1205e-125">Word - 320x531 pixels</span></span>
- <span data-ttu-id="1205e-126">Outlook — 348x535 пикселей</span><span class="sxs-lookup"><span data-stu-id="1205e-126">Outlook - 348x535 pixels</span></span>

<br/>

<span data-ttu-id="1205e-127">*Рис. 3. Office размеров области задач*</span><span class="sxs-lookup"><span data-stu-id="1205e-127">*Figure 3. Office task pane sizes*</span></span>

![Схема с отображением размеров области задач с разрешением 1366x768](../images/office-365-taskpane-sizes.png)

- <span data-ttu-id="1205e-129">Excel — 350x378 пикселей</span><span class="sxs-lookup"><span data-stu-id="1205e-129">Excel - 350x378 pixels</span></span>
- <span data-ttu-id="1205e-130">PowerPoint — 348x391 пикселей</span><span class="sxs-lookup"><span data-stu-id="1205e-130">PowerPoint - 348x391 pixels</span></span>
- <span data-ttu-id="1205e-131">Word — 329x445 пикселей</span><span class="sxs-lookup"><span data-stu-id="1205e-131">Word - 329x445 pixels</span></span>
- <span data-ttu-id="1205e-132">Outlook (в Интернете) — 320x570 пикселей</span><span class="sxs-lookup"><span data-stu-id="1205e-132">Outlook (on the web) - 320x570 pixels</span></span>

## <a name="personality-menu"></a><span data-ttu-id="1205e-133">Меню личных данных</span><span class="sxs-lookup"><span data-stu-id="1205e-133">Personality menu</span></span>

<span data-ttu-id="1205e-p103">Меню личных данных могут перекрывать элементы навигации и управления, расположенные в правой верхней части надстройки. Ниже указаны текущие размеры меню личных данных в Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="1205e-p103">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="1205e-136">Меню личных данных в Windows имеет размер 12 x 32 пикселей, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="1205e-136">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="1205e-137">*Рис. 4. Меню личных данных в Windows*</span><span class="sxs-lookup"><span data-stu-id="1205e-137">*Figure 4. Personality menu on Windows*</span></span>

![Схема, показывающая меню личности на Windows рабочем столе](../images/personality-menu-win.png)

<span data-ttu-id="1205e-139">В Mac меню личных данных имеет размер 26 x 26 пикселей, но сдвинуто на 8 пикселей влево и на 6 вниз, из-за чего оно занимает пространство размером 34 x 32 пикселя, как показано на изображении.</span><span class="sxs-lookup"><span data-stu-id="1205e-139">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="1205e-140">*Рис. 5. Меню личных данных на Mac*</span><span class="sxs-lookup"><span data-stu-id="1205e-140">*Figure 5. Personality menu on Mac*</span></span>

![Схема, показывающая меню личности на рабочем столе Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="1205e-142">Реализация</span><span class="sxs-lookup"><span data-stu-id="1205e-142">Implementation</span></span>

<span data-ttu-id="1205e-143">Ознакомьтесь с реализацией области задач на примере [надстройки Excel "Тенденции расходов банка WoodGrove" на JS](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="1205e-143">For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub.</span></span>

## <a name="see-also"></a><span data-ttu-id="1205e-144">См. также</span><span class="sxs-lookup"><span data-stu-id="1205e-144">See also</span></span>

- [<span data-ttu-id="1205e-145">Fabric Core в Office надстройки</span><span class="sxs-lookup"><span data-stu-id="1205e-145">Fabric Core in Office Add-ins</span></span>](fabric-core.md)
- [<span data-ttu-id="1205e-146">Конструктивные шаблоны для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="1205e-146">UX design patterns for Office Add-ins</span></span>](../design/ux-design-pattern-templates.md)
