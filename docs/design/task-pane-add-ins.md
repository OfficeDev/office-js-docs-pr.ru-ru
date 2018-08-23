---
title: Области задач в надстройках Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d60af0a31b9f96be17aa55bda789d13c386d0a5b
ms.sourcegitcommit: 7ecc1dc24bf7488b53117d7a83ad60e952a6f7aa
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/23/2018
ms.locfileid: "19437201"
---
# <a name="task-panes-in-office-add-ins"></a><span data-ttu-id="632eb-102">Области задач в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="632eb-102">Task panes in Office Add-ins</span></span>
 
<span data-ttu-id="632eb-p101">Области задач — это области интерфейса, которые обычно отображаются в правой части окна Word, PowerPoint, Excel и Outlook. Элементы области задач выполняют код для изменения документов или писем, а также для отображения данных. Используйте области задач, когда вам не нужно внедрять функции прямо в документ.</span><span class="sxs-lookup"><span data-stu-id="632eb-p101">Task panes are interface surfaces that typically appear on the right side of the window within Word, PowerPoint, Excel, and Outlook. Task panes give users access to interface controls that run code to modify documents or emails, or display data from a data source. Use task panes when you don't need to embed functionality directly into the document.</span></span>

<span data-ttu-id="632eb-106">*Рис. 1. Типичный макет области задач*</span><span class="sxs-lookup"><span data-stu-id="632eb-106">*Figure 1. Typical task pane layout*</span></span>

![Типичный макет области задач](../images/overview-with-app-task-pane.png)

## <a name="best-practices"></a><span data-ttu-id="632eb-108">Советы и рекомендации</span><span class="sxs-lookup"><span data-stu-id="632eb-108">Best practices</span></span>

|<span data-ttu-id="632eb-109">**Рекомендуется**</span><span class="sxs-lookup"><span data-stu-id="632eb-109">**Do**</span></span>|<span data-ttu-id="632eb-110">**Не рекомендуется**</span><span class="sxs-lookup"><span data-stu-id="632eb-110">**Don't**</span></span>|
|:-----|:--------|
|<ul><li><span data-ttu-id="632eb-111">Включите имя надстройки в название.</span><span class="sxs-lookup"><span data-stu-id="632eb-111">Include the name of your add-in in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="632eb-112">Не включайте в него название вашей компании.</span><span class="sxs-lookup"><span data-stu-id="632eb-112">Don't append your company name to the title.</span></span></li></ul>|
|<ul><li><span data-ttu-id="632eb-113">Используйте короткие описательные имена в названии.</span><span class="sxs-lookup"><span data-stu-id="632eb-113">Use short descriptive names in the title.</span></span></li></ul>|<ul><li><span data-ttu-id="632eb-114">Не включайте такие строки, как "Надстройка", "Для Word" или "для Office", в название надстройки.</span><span class="sxs-lookup"><span data-stu-id="632eb-114">Don't append strings such as “Add-in,” “For Word,” or “for Office” to the title of your add-in.</span></span></li></ul>|
|<ul><li><span data-ttu-id="632eb-115">Включите элемент навигации или управления, такой как CommandBar или Pivot, в верхнюю часть надстройки.</span><span class="sxs-lookup"><span data-stu-id="632eb-115">Include some navigational or commanding element such as the CommandBar or Pivot at the top of your add-in.</span></span></li></ul>||
|<ul><li><span data-ttu-id="632eb-116">Включите элемент фирменной символики, такой как BrandBar, в нижнюю часть надстройки, если только она не будет использоваться исключительно в Outlook.</span><span class="sxs-lookup"><span data-stu-id="632eb-116">Include a branding element such as the BrandBar at the bottom of your add-in unless your add-in is to be used within Outlook.</span></span></li></ul>||


## <a name="variants"></a><span data-ttu-id="632eb-117">Варианты</span><span class="sxs-lookup"><span data-stu-id="632eb-117">Variants</span></span>

<span data-ttu-id="632eb-p102">На изображениях ниже приведены области задач разных размеров с лентой Office при разрешении 1366 x 768. Чтобы вставить строку формул в Excel, требуется дополнительное пространство по вертикали.</span><span class="sxs-lookup"><span data-stu-id="632eb-p102">The following images show the various task pane sizes with the Office ribbon at a 1366x768 resolution. For Excel, additional vertical space is required to accommodate the formula bar.</span></span>  

<span data-ttu-id="632eb-120">*Рис. 2. Размеры области задач в классических приложениях Office 2016*</span><span class="sxs-lookup"><span data-stu-id="632eb-120">*Figure 2. Office 2016 desktop task pane sizes*</span></span>

![Размеры области задач в классических приложениях при разрешении 1366 x 768](../images/add-in-taskpane-sizes-desktop.png)

- <span data-ttu-id="632eb-122">Excel — 320 x 455</span><span class="sxs-lookup"><span data-stu-id="632eb-122">Excel - 320x455</span></span>
- <span data-ttu-id="632eb-123">PowerPoint — 320 x 531</span><span class="sxs-lookup"><span data-stu-id="632eb-123">PowerPoint - 320x531</span></span>
- <span data-ttu-id="632eb-124">Word — 320 x 531</span><span class="sxs-lookup"><span data-stu-id="632eb-124">Word - 320x531</span></span>
- <span data-ttu-id="632eb-125">Outlook — 348 x 535</span><span class="sxs-lookup"><span data-stu-id="632eb-125">Outlook - 348x535</span></span>

<br/>

<span data-ttu-id="632eb-126">*Рис. 3. Размеры области задач в Office 365*</span><span class="sxs-lookup"><span data-stu-id="632eb-126">*Figure 3. Office 365 task pane sizes*</span></span>

![Размеры области задач в классических приложениях при разрешении 1366 x 768](../images/add-in-taskpane-sizes-online.png)

- <span data-ttu-id="632eb-128">Excel — 350 x 378</span><span class="sxs-lookup"><span data-stu-id="632eb-128">Excel - 350x378</span></span>
- <span data-ttu-id="632eb-129">PowerPoint — 348 x 391</span><span class="sxs-lookup"><span data-stu-id="632eb-129">PowerPoint - 348x391</span></span>
- <span data-ttu-id="632eb-130">Word — 329 x 445</span><span class="sxs-lookup"><span data-stu-id="632eb-130">Word - 329x445</span></span>
- <span data-ttu-id="632eb-131">Outlook Web App — 320 x 570</span><span class="sxs-lookup"><span data-stu-id="632eb-131">Outlook Web App - 320x570</span></span>

## <a name="personality-menu"></a><span data-ttu-id="632eb-132">Меню личных данных</span><span class="sxs-lookup"><span data-stu-id="632eb-132">Personality menu</span></span>

<span data-ttu-id="632eb-p103">Меню личных данных могут перекрывать элементы навигации и управления, расположенные в правой верхней части надстройки. Ниже указаны текущие размеры меню личных данных в Windows и Mac.</span><span class="sxs-lookup"><span data-stu-id="632eb-p103">Personality menus can obstruct navigational and commanding elements located near the top right of the add-in. The following are the current dimensions of the personality menu on Windows and Mac.</span></span>

<span data-ttu-id="632eb-135">Меню личных данных в Windows имеет размер 12 x 32 пикселей, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="632eb-135">For Windows, the personality menu measures 12x32 pixels, as shown.</span></span>

<span data-ttu-id="632eb-136">*Рис. 4. Меню личных данных в Windows*</span><span class="sxs-lookup"><span data-stu-id="632eb-136">*Figure 4. Personality menu on Windows*</span></span>

![Меню личных данных на компьютере с Windows](../images/personality-menu-win.png)

<span data-ttu-id="632eb-138">В Mac меню личных данных имеет размер 26 x 26 пикселей, но сдвинуто на 8 пикселей влево и на 6 вниз, из-за чего оно занимает пространство размером 34 x 32 пикселя, как показано на изображении.</span><span class="sxs-lookup"><span data-stu-id="632eb-138">For Mac, the personality menu measures 26x26 pixels, but floats 8 pixels in from the right and 6 pixels from the top, which increases the space to 34x32 pixels, as shown.</span></span>

<span data-ttu-id="632eb-139">*Рис. 5. Меню личных данных на Mac*</span><span class="sxs-lookup"><span data-stu-id="632eb-139">*Figure 5. Personality menu on Mac*</span></span>

![Меню личных данных на Mac](../images/personality-menu-mac.png)

## <a name="implementation"></a><span data-ttu-id="632eb-141">Реализация</span><span class="sxs-lookup"><span data-stu-id="632eb-141">Implementation</span></span>

<span data-ttu-id="632eb-142">Ознакомьтесь с реализацией области задач на примере [надстройки Excel "Тенденции расходов банка WoodGrove" на JS](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="632eb-142">For a sample that implements a task pane, see [Excel Add-in JS WoodGrove Expense Trends](https://github.com/OfficeDev/Excel-Add-in-WoodGrove-Expense-Trends) on GitHub.</span></span> 


## <a name="see-also"></a><span data-ttu-id="632eb-143">См. также</span><span class="sxs-lookup"><span data-stu-id="632eb-143">See also</span></span>

- [<span data-ttu-id="632eb-144">Office UI Fabric в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="632eb-144">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md) 
- [<span data-ttu-id="632eb-145">Конструктивные шаблоны для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="632eb-145">UX design patterns for Office Add-ins</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)


